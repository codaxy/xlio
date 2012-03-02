using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class CellData
    {
        public object Value { get; set; }

        String formula;
        SharedFormula sharedFormula;

        public String Formula
        {
            get { return formula; }
            set { formula = value; sharedFormula = null; }
        }

        public SharedFormula SharedFormula
        {
            get { return sharedFormula; }
            set { sharedFormula = value; formula = null; }
        }

        internal CellStyle style;
        public CellStyle Style
        {
            get { return style ?? (style = new CellStyle()); }
            set { style = value; }
        }

        Range mergedCells;        
        public Range MergedRange
        {
            get { return mergedCells; }
            internal set { mergedCells = value; }
        }

        public bool IsMerged
        {
            get { return mergedCells != null; }

        }       
    }
}
