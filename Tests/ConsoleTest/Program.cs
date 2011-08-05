using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio;
using System.Diagnostics;
using System.Threading;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            System.IO.Directory.CreateDirectory(".\\write");
            foreach (var f in System.IO.Directory.GetFileSystemEntries(".", "read\\*.xlsx").OrderBy(a=>a))
            {
                try
                {
                    Console.Write("Reading file '{0}'...", f);
                    stopwatch.Reset(); stopwatch.Start();
                    Workbook wb = Workbook.ReadFile(f);
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

            Thread.Sleep(5000);

            
        }
    }
}
