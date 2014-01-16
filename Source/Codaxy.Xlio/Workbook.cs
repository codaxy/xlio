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
            DefaultFont = new CellFont();
            DefinedNames = new DefinedNameCollection();
        }

        public SheetCollection Sheets { get; private set; }
        public CellFont DefaultFont { get; set; }
        public DefinedNameCollection DefinedNames { get; set; }
    }
}
