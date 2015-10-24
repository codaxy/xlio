using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio;
using System.Diagnostics;
using System.Threading;
using Codaxy.Xlio.Samples.Usage;

namespace Codaxy.Xlio.Samples
{
    class Program
    {
        class A
        {
            public String Name
            {
                get; set;
            }
        }
        static void Main(string[] args)
        {
            PetaTest.Runner.RunMain(args);
        }


    }
}
