using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace XsdDir
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo dir;
            String argss;
            if (args.Length > 0) {
                dir = new DirectoryInfo(args[0]);
                argss =  String.Join(" ", args, 1, args.Length-1);
            } else {
                dir = new DirectoryInfo(Directory.GetCurrentDirectory());
                argss = "";
            }
            var files = String.Join(" ", dir.GetFiles("*.xsd").Select(a => a.FullName).ToArray());           
            Console.WriteLine("Calling xsd with arguments: {0}", String.Join(" ", args.Skip(1).ToArray())); 
            using (var p = new Process())
            {
                p.StartInfo = new ProcessStartInfo(@"C:\Program Files\Microsoft SDKs\Windows\v6.0A\bin\xsd.exe")
                {
                    Arguments = files + " " + argss,
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                
                p.Start();
                p.WaitForExit();
                var err = p.StandardError.ReadToEnd();
                if (!String.IsNullOrEmpty(err))
                    Console.Error.Write(err);
                Console.Write(p.StandardOutput.ReadToEnd());
            }
        }
    }
}
