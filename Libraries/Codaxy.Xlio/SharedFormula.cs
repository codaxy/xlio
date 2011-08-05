using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class SharedFormula
    {
        public String Formula { get; set; }
        public Range Range { get; set; }
        public Cell Origin { get; set; }

        public SharedFormula(String formula, Range range, Cell origin)
        {
            this.Formula = formula;
            this.Range = range;
            this.Origin = origin;
        }

        public String GetCellFormula(Cell cell)
        {
            throw new NotImplementedException();
        }
    }
}
