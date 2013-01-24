using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    /// <summary>
    /// CellData contains information about the cell, such as the cell value, formula, formatting, etc.
    /// </summary>
    public class CellData
    {
        /// <summary>
        /// Value.
        /// </summary>
        public object Value { get; set; }

        String formula;
        SharedFormula sharedFormula;

        /// <summary>
        /// Formula used to calculate the cell value.
        /// </summary>
        public String Formula
        {
            get { return formula; }
            set { formula = value; sharedFormula = null; }
        }

        /// <summary>
        /// A formula that can be shared between multiple cells. Use it to assign same formula to a range of cells.
        /// </summary>
        public SharedFormula SharedFormula
        {
            get { return sharedFormula; }
            set { sharedFormula = value; formula = null; }
        }

        internal CellStyle style;

        /// <summary>
        /// Contains cell style information.
        /// </summary>
        public CellStyle Style
        {
            get { return style ?? (style = new CellStyle()); }
            set { style = value; }
        }

        Range mergedCells;   
     
        /// <summary>
        /// Contains information about the marged area this cell belongs to.
        /// </summary>
        public Range MergedRange
        {
            get { return mergedCells; }
            internal set { mergedCells = value; }
        }

        /// <summary>
        /// True if cell is merged.
        /// </summary>
        public bool IsMerged
        {
            get { return mergedCells != null; }
        }       
    }
}
