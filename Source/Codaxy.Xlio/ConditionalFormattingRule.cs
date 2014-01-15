using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class ConditionalFormattingRule
    {
        public ConditionalFormattingRule()
        {
            Formulas = new List<string>();
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
        public List<string> Formulas { get; set; }

        public ConditionalFormattingColorScale ColorScale { get; set; }

        public ConditionalFormattingDataBar DataBar { get; set; }

        public ConditionalFormattingIconSet IconSet { get; set; }

        public ConditionalFormattingType Type { get; set; }

        public ConditionalFormattingOperator Operator { get; set; }
    }

    public enum ConditionalFormattingType
    {
        cellIs,
        colorScale,
        dataBar,
        iconSet
    }

    public enum ConditionalFormattingOperator
    {
        lessThan,
        lessThanOrEqual,
        equal,
        notEqual,
        greaterThanOrEqual,
        greaterThan,
        between,
        notBetween,
        containsText,
        notContains,
        beginsWith,
        endsWith
    }

    public enum CFVOType
    {
        num,
        percent,
        max,
        min,
        formula,
        percentile,
    }

}
