using System;
using System.Collections.Generic;
using Codaxy.Xlio.IO;

namespace Codaxy.Xlio
{
    //SRP: Excel workbook model
    public partial class Workbook
    {
        public Workbook()
        {
            Sheets = new SheetCollection();
        }

        public SheetCollection Sheets { get; private set; }        
    }
}
