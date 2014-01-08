using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PetaTest;
using System.Diagnostics;

namespace Codaxy.Xlio
{
    [TestFixture]
    class ReadWriteTests 
    {
        [Test]
        public void ReadWriteTest()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            System.IO.Directory.CreateDirectory(".\\write");
            foreach (var f in System.IO.Directory.GetFileSystemEntries(".", "read\\*.xlsx").OrderBy(a => a))
            {
                try
                {
                    Console.Write("Reading file '{0}'...", f);
                    stopwatch.Reset(); stopwatch.Start();
                    Workbook wb = Workbook.Load(f);
                    Console.WriteLine("{0:0.00 ms}", stopwatch.Elapsed.TotalMilliseconds);
                    stopwatch.Reset(); stopwatch.Start();
                    var of = f.Replace("\\read\\", "\\write\\");
                    Console.Write("Writing file '{0}'...", of);
                    wb.Save(of);
                    Console.WriteLine("{0:0.00 ms}", stopwatch.Elapsed.TotalMilliseconds);
                }
                catch (Exception ex)
                {
                    Console.Write(ex.Message);
                    Console.WriteLine(ex.StackTrace);
                    Console.ReadKey();
                }
            }
        }

        [Test]
        public void BoolTest()
        {
            var f = new Workbook();
            var sheet = new Sheet("1");
            f.Sheets.AddSheet(sheet);
            sheet[0, 0].Value = true;
            f.Save("bool.xlsx");

            f = Workbook.Load("bool.xlsx");
            Assert.AreEqual(true, f.Sheets[0][0, 0].Value);
        }
    }
}
