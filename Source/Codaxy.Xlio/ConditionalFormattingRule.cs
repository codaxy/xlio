using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public abstract class ConditionalFormattingRule
    {
        public ConditionalFormattingRule()
        {
            Priority = 1;
            Type = ConditionalFormattingType.CellIs;
        }
        /*
        private string[] formulaField;

        private CT_ColorScale colorScaleField;

        private CT_DataBar dataBarField;

        private CT_IconSet iconSetField;

        private CT_ExtensionList extLstField;

        private ST_CfType typeField;
         
        private ST_ConditionalFormattingOperator operatorField;
        */

        /*
        public List<string> Formulas { get; set; }

        public ConditionalFormattingColorScale ColorScale { get; set; }

        public ConditionalFormattingDataBar DataBar { get; set; }

        public ConditionalFormattingIconSet IconSet { get; set; }

        public ConditionalFormattingOperator Operator { get; set; }
         
        */
        public ConditionalFormattingType Type { get; set; }
        public int Priority { get; set; }
    }

    public enum ConditionalFormattingType
    {
        /*
        CellIs,
        ColorScale,
        DataBar,
        IconSet,
        BeginsWith //ADDED AFTEER
         * 
         * */
        Expression,
        CellIs,
        ColorScale,
        DataBar,
        IconSet,
        Top10,
        UniqueValues,
        DuplicateValues,
        ContainsText,
        NotContainsText,
        BeginsWith,
        EndsWith,
        ContainsBlanks,
        NotContainsBlanks,
        ContainsErrors,
        NotContainsErrors,
        TimePeriod,
        AboveAverage
    }

    public enum ConditionalFormattingOperator
    {
        LessThan,
        LessThanOrEqual,
        Equal,
        NotEqual,
        GreaterThanOrEqual,
        GreaterThan,

        Between,
        NotBetween,

        ContainsText,
        NotContains,

        BeginsWith,
        EndsWith
    }

    public enum CFVOType
    {
        Num,
        Percent,
        Max,
        Min,
        Formula,
        Percentile,
    }

}
