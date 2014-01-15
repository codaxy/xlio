using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class ConditionalFormatting
    {
        public ConditionalFormatting()
        {
            Rules = new List<ConditionalFormattingRule>();
            SequenceOfReferences = new List<string>();
        }

        public List<ConditionalFormattingRule> Rules { get; set; }

        public List<string> SequenceOfReferences { get; set; }
    }

    public class ConditionalFormattingValueObject
    {
        public CFVOType Type { get; set; }
        public string Value { get; set; }
    }

    public class ConditionalFormattingColorScale
    {

        public ConditionalFormattingColorScale()
        {
            CFVOList = new List<ConditionalFormattingValueObject>();
            Colors = new List<Color>();
        }
        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public List<Color> Colors { get; set; }
    }

    public class ConditionalFormattingDataBar
    {

        public ConditionalFormattingDataBar()
        {
            CFVOList = new List<ConditionalFormattingValueObject>();
        }
        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public Color Color { get; set; }
    }

    public class ConditionalFormattingIconSet
    {
        /*
        private CT_Cfvo[] cfvoField
        private ST_IconSetType iconSetField
        private bool showValueField
        private bool percentField
        private bool reverseField
        */
        public ConditionalFormattingIconSet(ConditionalFormattingTypeIconSetType IconSetType = ConditionalFormattingTypeIconSetType.Item3Arrows, bool showValue = true)
        {
            CFVOList = new List<ConditionalFormattingValueObject>();
            ShowValue = showValue;
        }

        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public ConditionalFormattingTypeIconSetType IconSetType { get; set; }
        public bool ShowValue { get; set; }
    }

    public enum ConditionalFormattingTypeIconSetType
    {
        Item3Arrows,
        Item3ArrowsGray,
        Item3Flags,
        Item3TrafficLights1,
        Item3TrafficLights2,
        Item3Signs,
        Item3Symbols,
        Item3Symbols2,

        Item4Arrows,
        Item4ArrowsGray,
        Item4RedToBlack,
        Item4Rating,
        Item4TrafficLights,

        Item5Arrows,
        Item5ArrowsGray,
        Item5Rating,
        Item5Quarters
    }
}
