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

        public string GetFirstCellStringValue() 
        {
            if (Ranges.Count == 0)
                return null;
            string firstRange = Ranges[0];
            if (!firstRange.Contains(":"))
                return firstRange;
            int position = firstRange.IndexOf(":");
            string firstCell = firstRange.Substring(0, position);
            return firstCell;
        }

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
        public IconSet(ConditionalFormattingIconSetType iconSetType = ConditionalFormattingIconSetType.Item3Arrows, bool showValue = true) : base()
        {
            base.Type = ConditionalFormattingType.IconSet;
            IconSetType = iconSetType;
            CFVOList = GenerateCFVOs();
            ShowValue = showValue;
        }

        public List<ConditionalFormattingValueObject> CFVOList { get; set; }
        public ConditionalFormattingIconSetType IconSetType { get; set; }
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
            ConditionalFormattingIconSetType[] fourIcons = { 
                ConditionalFormattingIconSetType.Item4Arrows,
                ConditionalFormattingIconSetType.Item4ArrowsGray,
                ConditionalFormattingIconSetType.Item4RedToBlack,
                ConditionalFormattingIconSetType.Item4Rating,
                ConditionalFormattingIconSetType.Item4TrafficLights
            };
            ConditionalFormattingIconSetType[] fiveIcons = { 
                ConditionalFormattingIconSetType.Item5Arrows,
                ConditionalFormattingIconSetType.Item5ArrowsGray,
                ConditionalFormattingIconSetType.Item5Rating,
                ConditionalFormattingIconSetType.Item5Quarters
            };
            if (fourIcons.Contains(IconSetType))
                return 4;
            if (fiveIcons.Contains(IconSetType))
                return 5;
            return 3;
        }
    }

    public enum ConditionalFormattingIconSetType
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
            IsBottom = false;
            Rank = 10;
            IsPercent = false;
            StdDev = 0;
            IsAboveAverage = true;
            IsEqualAverage = false;
            IsStdDev = false;
        }
        public uint dxfId { get; set; }
        public uint Rank { get; set; }
        public bool IsBottom { get; set; }
        public bool IsPercent { get; set; }
        public bool IsAboveAverage { get; set; }
        public bool IsEqualAverage { get; set; }
        public bool IsStdDev { get; set; }
        public int StdDev { get; set; }
        public string Text { get; set; }
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

        //tekst operacije za sada ne rade, jer koriste formule
        /*
        public static ConditionalFormattingCondition Contains(string value, CellStyle style)
        {
            ConditionalFormattingCondition cfc = GetConditionForOneValue(value, style, ConditionalFormattingOperator.ContainsText);
            cfc.Type = ConditionalFormattingType.ContainsText;
            return cfc;
        }

        public static ConditionalFormattingCondition NotContains(string value, CellStyle style)
        {
            ConditionalFormattingCondition cfc = GetConditionForOneValue(value, style, ConditionalFormattingOperator.NotContains);
            cfc.Type = ConditionalFormattingType.NotContainsText;
            return cfc;
        }

        public static ConditionalFormattingCondition BeginsWith(string value, CellStyle style)
        {
            ConditionalFormattingCondition cfc = GetConditionForOneValue(value, style, ConditionalFormattingOperator.BeginsWith);
            cfc.Type = ConditionalFormattingType.BeginsWith;
            return cfc;
        }

        public static ConditionalFormattingCondition EndsWith(string value, CellStyle style)
        {
            ConditionalFormattingCondition cfc = GetConditionForOneValue(value, style, ConditionalFormattingOperator.EndsWith);
            cfc.Type = ConditionalFormattingType.EndsWith;
            return cfc;
        }
         * */
        public static ConditionalFormattingCondition Top(uint rank, CellStyle style, bool percent = false)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition(); 
            cfc.Type = ConditionalFormattingType.Top10;
            cfc.Rank = rank;
            cfc.IsPercent = percent;
            cfc.Style = style;
            return cfc;
        }
        public static ConditionalFormattingCondition Bottom(uint rank, CellStyle style, bool percent = false)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.Top10;
            cfc.Rank = rank;
            cfc.IsBottom = true;
            cfc.IsPercent = percent;
            cfc.Style = style;
            return cfc;
        }
        public static ConditionalFormattingCondition AboveAverage(CellStyle style, bool equal = false)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.AboveAverage;
            cfc.Style = style;
            cfc.IsAboveAverage = true;
            cfc.IsEqualAverage = equal;
            return cfc;
        }
        public static ConditionalFormattingCondition BelowAverage(CellStyle style, bool equal = false)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.AboveAverage;
            cfc.Style = style;
            cfc.IsAboveAverage = false;
            cfc.IsEqualAverage = equal;
            return cfc;
        }
        public static ConditionalFormattingCondition StandardDeviationAbove(CellStyle style, int stdDev = 1)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.AboveAverage;
            cfc.Style = style;
            cfc.IsAboveAverage = true;
            cfc.IsEqualAverage = false;
            cfc.IsStdDev = true;
            cfc.StdDev = stdDev;
            return cfc;
        }
        public static ConditionalFormattingCondition StandardDeviationBelow(CellStyle style, int stdDev = 1)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.AboveAverage;
            cfc.Style = style;
            cfc.IsAboveAverage = false;
            cfc.IsEqualAverage = false;
            cfc.IsStdDev = true;
            cfc.StdDev = stdDev;
            return cfc;
        }
        public static ConditionalFormattingCondition UniqueValues(CellStyle style)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.UniqueValues;
            cfc.Style = style;
            return cfc;
        }
        public static ConditionalFormattingCondition DuplicateValues(CellStyle style)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.DuplicateValues;
            cfc.Style = style;
            return cfc;
        }
        public static ConditionalFormattingCondition Expression(string formula, CellStyle style)
        {
            ConditionalFormattingCondition cfc = new ConditionalFormattingCondition();
            cfc.Type = ConditionalFormattingType.Expression;
            cfc.Formula1 = formula;
            cfc.Style = style;
            return cfc;
        }
    }
}
