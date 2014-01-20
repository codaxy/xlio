using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{

    public class ConditionalFormatting : IEnumerable<ConditionalFormattingRule>
    {
        public ConditionalFormatting(Range range = null)
        {
            Rules = new List<ConditionalFormattingRule>();
            Ranges = new List<string>();
            if (range != null)
                AddRange(range);
        }

        public List<ConditionalFormattingRule> Rules { get; set; }

        public List<string> Ranges { get; set; }

        public void AddRange(Range range) 
        {
            Ranges.Add(Range.Format(range));
        }

        public void AddCell(Cell cell)
        {
            Ranges.Add(Cell.Format(cell.Row, cell.Col));
        }

        public ConditionalFormattingRule AddRule(ConditionalFormattingRule rule) 
        {
            Rules.Add(rule);
            return rule;
        }

        public void Remove(ConditionalFormattingRule rule)
        {
            Rules.Remove(rule);
        }

        public ConditionalFormattingRule this[int index]
        {
            get
            {
                return Rules[index];
            }
            set
            {
                Rules[index] = value;
            }
        }

        int Count { get { return Rules.Count; } }

        public IEnumerator<ConditionalFormattingRule> GetEnumerator()
        {
            return Rules.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return Rules.GetEnumerator();
        }

    }
  
    public class ConditionalFormattingValueObject
    {
        public ConditionalFormattingValueObject(){}
        public ConditionalFormattingValueObject(CFVOType type, string value)
        {
            Type = type; 
            Value = value;
        }
        public CFVOType Type { get; set; }
        public string Value { get; set; }
    }

    //COLOR SCALE
    public class ColorScale : ConditionalFormattingRule
    {
        public ColorScale(ConditionalFormattingColorScaleType type = ConditionalFormattingColorScaleType.TwoColorScale)
            : base()
        {
            base.Type = ConditionalFormattingType.ColorScale;
            Color minimum = Color.Red, maximum = Color.Green, midpoint = Color.Yellow;
            CFVOList = new List<ConditionalFormattingValueObject>();
            Colors = new List<Color>();
            if (type == ConditionalFormattingColorScaleType.TwoColorScale)
                SetTwoColorScale(minimum, maximum);
            else
                SetThreeColorScale(minimum, midpoint, maximum);
        }

        private void SetTwoColorScale(Color minimum, Color maximum)
        {
            var cfvoMin = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Min
            };
            var cfvoMax = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Max
            };
            CFVOList.Add(cfvoMin);
            CFVOList.Add(cfvoMax);
            Colors.Add(minimum);
            Colors.Add(maximum);
        }

        private void SetThreeColorScale(Color minimum, Color midpoint, Color maximum, int centerPercentage = 50)
        {
            var cfvoMin = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Min
            };
            var cfvoMid = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Percent,
                Value = centerPercentage.ToString()
            };
            var cfvoMax = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Max
            };
            CFVOList.Add(cfvoMin);
            CFVOList.Add(cfvoMid);
            CFVOList.Add(cfvoMax);
            Colors.Add(minimum);
            Colors.Add(midpoint);
            Colors.Add(maximum);
        }

        public static ColorScale TwoColor(Color minimum, Color maximum)
        {
            ColorScale cs = new ColorScale(ConditionalFormattingColorScaleType.TwoColorScale);
            cs.SetTwoColorScale(minimum, maximum);
            return cs;    
        }

        public static ColorScale ThreeColor(Color minimum, Color midpoint, Color maximum, int centerPercentage = 50)
        {
            ColorScale cs = new ColorScale(ConditionalFormattingColorScaleType.ThreeColorScale);
            cs.SetThreeColorScale(minimum, midpoint, maximum, centerPercentage);
            return cs;
        }

        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public List<Color> Colors { get; set; }
    }

    public enum ConditionalFormattingColorScaleType
    { 
        TwoColorScale,
        ThreeColorScale
    }

    //DATA BAR
    public class DataBar : ConditionalFormattingRule
    {
        public DataBar(Color color)
            : base()
        {
            base.Type = ConditionalFormattingType.DataBar;
            CFVOList = new List<ConditionalFormattingValueObject>();
            Color = color;
            SetDataBarValues();
        }
        public DataBar() : base()
        {
            base.Type = ConditionalFormattingType.DataBar;
            CFVOList = new List<ConditionalFormattingValueObject>();
            Color = Color.Blue;
            SetDataBarValues();
        }
        private void SetDataBarValues()
        {
            var cfvoMin = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Min
            };
            var cfvoMax = new ConditionalFormattingValueObject
            {
                Type = CFVOType.Max
            };
            CFVOList.Add(cfvoMin);
            CFVOList.Add(cfvoMax);
        }

        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public Color Color { get; set; }
    }

    //ICON SET
    public class IconSet : ConditionalFormattingRule
    {
        /*
        private CT_Cfvo[] cfvoField
        private ST_IconSetType iconSetField
        private bool showValueField
        private bool percentField
        private bool reverseField
        */
        public IconSet(ConditionalFormattingTypeIconSetType iconSetType = ConditionalFormattingTypeIconSetType.Item3Arrows, bool showValue = true) : base()
        {
            base.Type = ConditionalFormattingType.IconSet;
            IconSetType = iconSetType;
            CFVOList = GenerateCFVOs();
            ShowValue = showValue;
        }

        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public ConditionalFormattingTypeIconSetType IconSetType { get; set; }
        public bool ShowValue { get; set; }

        private List<ConditionalFormattingValueObject> GenerateCFVOs() 
        {
            int numberOfIcons = GetNumberOfIcons();
            decimal step = 100.0m / numberOfIcons;
            List<ConditionalFormattingValueObject> list = new List<ConditionalFormattingValueObject>();
            for (int i = 0; i < numberOfIcons; i++)
            {
                int percent = (int)Math.Round(i * step);
                ConditionalFormattingValueObject cfvo = new ConditionalFormattingValueObject
                {
                    Type = CFVOType.Percent,
                    Value = percent.ToString()
                };
                list.Add(cfvo);
            }
            return list;
        }

        private int GetNumberOfIcons() 
        {
            ConditionalFormattingTypeIconSetType[] fourIcons = { 
                ConditionalFormattingTypeIconSetType.Item4Arrows,
                ConditionalFormattingTypeIconSetType.Item4ArrowsGray,
                ConditionalFormattingTypeIconSetType.Item4RedToBlack,
                ConditionalFormattingTypeIconSetType.Item4Rating,
                ConditionalFormattingTypeIconSetType.Item4TrafficLights
            };
            ConditionalFormattingTypeIconSetType[] fiveIcons = { 
                ConditionalFormattingTypeIconSetType.Item5Arrows,
                ConditionalFormattingTypeIconSetType.Item5ArrowsGray,
                ConditionalFormattingTypeIconSetType.Item5Rating,
                ConditionalFormattingTypeIconSetType.Item5Quarters
            };
            if (fourIcons.Contains(IconSetType))
                return 4;
            if (fiveIcons.Contains(IconSetType))
                return 5;
            return 3;
        }
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

    public class ConditionalFormattingCondition : ConditionalFormattingRule
    {
        public ConditionalFormattingCondition(CellStyle style = null, ConditionalFormattingOperator @operator = ConditionalFormattingOperator.LessThan, 
            string formula1 = "", string formula2 = "", string formula3 = "") : base()
        {
            base.Type = ConditionalFormattingType.CellIs;
            Style = style;
            dxfId = 0;
            Operator = @operator;
            Formula1 = formula1;
            Formula2 = formula2;
            Formula3 = formula3;
        }
        public uint dxfId { get; set; }
        public CellStyle Style { get; set; }
        public ConditionalFormattingOperator Operator { get; set; }
        public string Formula1
        {
            get { return _Formula1; }
            set { _Formula1 = RemoveEqualChar(value); }
        }
        public string Formula2
        {
            get { return _Formula2; }
            set { _Formula2 = RemoveEqualChar(value); }
        }
        public string Formula3
        {
            get { return _Formula3; }
            set { _Formula3 = RemoveEqualChar(value); }
        }
        private string _Formula1 { get; set; }
        private string _Formula2 { get; set; }
        private string _Formula3 { get; set; }
        private string RemoveEqualChar(string formula) 
        {
            if (formula == null) return null;
            if (formula.StartsWith("="))
                return formula.Substring(1);
            return formula;
        }

        private static ConditionalFormattingCondition GetConditionForOneValue(string value, CellStyle style, ConditionalFormattingOperator @operator) 
        {
            return new ConditionalFormattingCondition(style, @operator, value);
        }

        private static ConditionalFormattingCondition GetConditionForTwoValues(string startValue, string endValue, CellStyle style, ConditionalFormattingOperator @operator)
        {
            return new ConditionalFormattingCondition(style, @operator, startValue, endValue);
        }

        public static ConditionalFormattingCondition LessThan(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.LessThan);
        }

        public static ConditionalFormattingCondition LessThanOrEqual(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.LessThanOrEqual);
        }

        public static ConditionalFormattingCondition Equal(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.Equal);
        }

        public static ConditionalFormattingCondition NotEqual(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.NotEqual);
        }

        public static ConditionalFormattingCondition GreaterThanOrEqual(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.GreaterThanOrEqual);
        }

        public static ConditionalFormattingCondition GreaterThan(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.GreaterThan);
        }

        public static ConditionalFormattingCondition Between(string startValue, string endValue, CellStyle style)
        {
            return GetConditionForTwoValues(startValue, endValue, style, ConditionalFormattingOperator.Between);
        }

        public static ConditionalFormattingCondition NotBetween(string startValue, string endValue, CellStyle style)
        {
            return GetConditionForTwoValues(startValue, endValue, style, ConditionalFormattingOperator.NotBetween);
        }

        public static ConditionalFormattingCondition ContainsText(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.ContainsText);
        }

        public static ConditionalFormattingCondition NotContains(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.NotContains);
        }

        public static ConditionalFormattingCondition BeginsWith(string value, CellStyle style)
        {
            //return GetConditionForOneValue(value, style, ConditionalFormattingOperator.BeginsWith);
            ConditionalFormattingCondition c = GetConditionForOneValue(value, style, ConditionalFormattingOperator.BeginsWith);
            c.Type = ConditionalFormattingType.BeginsWith;
            return c;
        }

        public static ConditionalFormattingCondition EndsWith(string value, CellStyle style)
        {
            return GetConditionForOneValue(value, style, ConditionalFormattingOperator.EndsWith);
        }
    }
}
