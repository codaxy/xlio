using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class ConditionalFormatting : IEnumerable<CFRule>
    {
        public ConditionalFormatting(Range range = null)
        {
            Rules = new List<CFRule>();
            Ranges = new List<string>();
            if (range != null)
                AddRange(range);
        }

        public List<CFRule> Rules { get; set; }
        public List<string> Ranges { get; set; }

        public string GetFirstCellStringValue() 
        {
            if (Ranges.Count == 0)
                throw new InvalidOperationException("At least one range should be specified.");
            
            string firstRange = Ranges[0];
            if (!firstRange.Contains(":"))
                return firstRange;
            int position = firstRange.IndexOf(":");
            string firstCell = firstRange.Substring(0, position);
            return firstCell;
        }

        public void Remove(CFRule rule)
        {
            Rules.Remove(rule);
        }

        public CFRule this[int index]
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

        public IEnumerator<CFRule> GetEnumerator()
        {
            return Rules.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return Rules.GetEnumerator();
        }

        public void AddRange(String range)
        {
            Ranges.Add(range);
        }

        public void AddRange(Range range) 
        {
            Ranges.Add(Range.Format(range));
        }

        public void AddCell(Cell cell)
        {
            Ranges.Add(Cell.Format(cell.Row, cell.Col));
        }

        public CFRule AddRule(CFRule rule) 
        {
            Rules.Add(rule);
            return rule;
        }

        #region NoDxfs
        /// <summary>
        /// Add created color scale.
        /// </summary>
        /// <param name="colorScale"></param>
        /// <returns></returns>
        public CFRule AddColorScale(ColorScale colorScale)
        {
            CFRule rule = new CFRule(CFType.ColorScale);
            rule.ColorScale = colorScale;
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add two color scale.
        /// </summary>
        /// <param name="colorScale"></param>
        /// <returns></returns>
        public CFRule AddColorScale(Color minimum, Color maximum)
        {
            CFRule rule = new CFRule(CFType.ColorScale);
            rule.ColorScale = ColorScale.TwoColor(minimum, maximum);
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add three color scale.
        /// </summary>
        /// <param name="minimum"></param>
        /// <param name="midpoint"></param>
        /// <param name="maximum"></param>
        /// <param name="centerPercentage"></param>
        /// <returns></returns>
        public CFRule AddColorScale(Color minimum, Color midpoint, Color maximum, int centerPercentage = 50)
        {
            CFRule rule = new CFRule(CFType.ColorScale);
            rule.ColorScale = ColorScale.TwoColor(minimum, maximum);
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add color scale with specified ConditionalFormattingValueObject and Color lists.
        /// </summary>
        /// <param name="cfvoList"></param>
        /// <param name="colors"></param>
        /// <returns></returns>
        public CFRule AddColorScale(List<CFVO> cfvoList, List<Color> colors)
        {
            CFRule rule = new CFRule(CFType.ColorScale);
            ColorScale cs = new ColorScale();
            cs.CFVOList = cfvoList;
            cs.Colors = colors;
            rule.ColorScale = cs;
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Default two color scale with red color as minimum and green as maxmimum.
        /// </summary>
        /// <returns></returns>
        public CFRule AddDefaultTwoColorScale()
        {
            CFRule rule = new CFRule(CFType.ColorScale);
            rule.ColorScale = ColorScale.DefaultTwoColor();
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Default three color scale with red color as minimum, yellow as midpoint and green as maximum.
        /// </summary>
        /// <param name="centerPercentage"></param>
        /// <returns></returns>
        public CFRule AddDefaultThreeColorScale(int centerPercentage = 50)
        {
            CFRule rule = new CFRule(CFType.ColorScale);
            rule.ColorScale = ColorScale.DefaultThreeColor(centerPercentage);
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add created data bar.
        /// </summary>
        /// <param name="dataBar"></param>
        /// <returns></returns>
        public CFRule AddDataBar(DataBar dataBar)
        {
            CFRule rule = new CFRule(CFType.DataBar);
            rule.DataBar = dataBar;
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add default data bar.
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public CFRule AddDefaultDataBar(Color color)
        {
            CFRule rule = new CFRule(CFType.DataBar);
            rule.DataBar = DataBar.CreateDefaultDataBar(color);
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add data bar with specified ConditionalFormattingValueObject and Color lists.
        /// </summary>
        /// <param name="cfvoList"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public CFRule AddDataBar(List<CFVO> cfvoList, Color color)
        {
            CFRule rule = new CFRule(CFType.DataBar);
            DataBar db = new DataBar();
            db.CFVOList = cfvoList;
            db.Color = color;
            rule.DataBar = db;
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add data bar with specified CFVOs for start and end values, as well as color.
        /// </summary>
        /// <param name="cfvoList"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public CFRule AddDataBar(CFVO start, CFVO end, Color color)
        {
            CFRule rule = new CFRule(CFType.DataBar);
            DataBar db = new DataBar();
            db.CFVOList = new List<CFVO>();
            db.CFVOList.Add(start);
            db.CFVOList.Add(end);
            db.Color = color;
            rule.DataBar = db;
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add created icon set.
        /// </summary>
        /// <param name="iconSet"></param>
        /// <returns></returns>
        public CFRule AddIconSet(IconSet iconSet)
        {
            CFRule rule = new CFRule(CFType.IconSet);
            rule.IconSet = iconSet;
            AddRule(rule);
            return rule;
        }

        /// <summary>
        /// Add default icon set.
        /// </summary>
        /// <param name="iconSetType"></param>
        /// <param name="showValue"></param>
        /// <returns></returns>
        public CFRule AddDefaultIconSet(IconSetType iconSetType = IconSetType.Item3Arrows, bool showValue = true)
        {
            CFRule rule = new CFRule(CFType.IconSet);
            rule.IconSet = IconSet.CreateDefaultIconSet(iconSetType, showValue);
            AddRule(rule);
            return rule;
        }
        #endregion

        #region Dxfs
        public CFRule AddTop(int rank, CellStyle style, bool percent = false)
        {
            CFRule rule = new CFRule(CFType.Top10);
            rule.Rank = rank;
            rule.IsPercent = percent;
            rule.Style = style;
            AddRule(rule);
            return rule;
        }
        public CFRule AddBottom(int rank, CellStyle style, bool percent = false)
        {
            CFRule rule = new CFRule(CFType.Top10);
            rule.Rank = rank;
            rule.IsBottom = true;
            rule.IsPercent = percent;
            rule.Style = style;
            AddRule(rule);
            return rule;
        }
        private CFRule GetRuleForOneValue(CFType type, string value, CellStyle style, CFOperator @operator, string text = null)
        {
            value = CFRule.CreateFormula(value, type, GetFirstCellStringValue());
            CFRule rule = new CFRule(type, style, @operator, value);
            rule.Text = text;
            AddRule(rule);
            return rule;
        }
        private CFRule GetRuleForTwoValues(CFType type, string startValue, string endValue, CellStyle style, CFOperator @operator)
        {
            CFRule rule = new CFRule(type, style, @operator, startValue, endValue);
            AddRule(rule);
            return rule;
        }
        public CFRule AddLessThan(string value, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForOneValue(type, value, style, CFOperator.LessThan);
        }

        public CFRule AddLessThanOrEqual(string value, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForOneValue(type, value, style, CFOperator.LessThanOrEqual);
        }

        public CFRule AddEqual(string value, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForOneValue(type, value, style, CFOperator.Equal);
        }

        public CFRule AddNotEqual(string value, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForOneValue(type, value, style, CFOperator.NotEqual);
        }

        public CFRule AddGreaterThanOrEqual(string value, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForOneValue(type, value, style, CFOperator.GreaterThanOrEqual);
        }

        public CFRule AddGreaterThan(string value, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForOneValue(type, value, style, CFOperator.GreaterThan);
        }

        public CFRule AddBetween(string startValue, string endValue, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForTwoValues(type, startValue, endValue, style, CFOperator.Between);
        }

        public CFRule AddNotBetween(string startValue, string endValue, CellStyle style)
        {
            CFType type = CFType.CellIs;
            return GetRuleForTwoValues(type, startValue, endValue, style, CFOperator.NotBetween);
        }

        public CFRule AddAboveAverage(CellStyle style, bool equal = false)
        {
            CFRule rule = new CFRule(CFType.AboveAverage);
            rule.Style = style;
            rule.IsAboveAverage = true;
            rule.IsEqualAverage = equal;
            AddRule(rule);
            return rule;
        }
        public CFRule AddBelowAverage(CellStyle style, bool equal = false)
        {
            CFRule rule = new CFRule(CFType.AboveAverage);
            rule.Style = style;
            rule.IsAboveAverage = false;
            rule.IsEqualAverage = equal;
            AddRule(rule);
            return rule;
        }
        public CFRule AddStandardDeviationAbove(CellStyle style, int stdDev = 1)
        {
            CFRule rule = new CFRule(CFType.AboveAverage);
            rule.Style = style;
            rule.IsAboveAverage = true;
            rule.IsEqualAverage = false;
            rule.IsStdDev = true;
            rule.StdDev = stdDev;
            AddRule(rule);
            return rule;
        }
        public CFRule AddStandardDeviationBelow(CellStyle style, int stdDev = 1)
        {
            CFRule rule = new CFRule(CFType.AboveAverage);
            rule.Style = style;
            rule.IsAboveAverage = false;
            rule.IsEqualAverage = false;
            rule.IsStdDev = true;
            rule.StdDev = stdDev;
            AddRule(rule);
            return rule;
        }
        public CFRule AddUniqueValues(CellStyle style)
        {
            CFRule rule = new CFRule(CFType.UniqueValues);
            rule.Style = style;
            AddRule(rule);
            return rule;
        }
        public CFRule AddDuplicateValues(CellStyle style)
        {
            CFRule rule = new CFRule(CFType.DuplicateValues);
            rule.Style = style;
            AddRule(rule);
            return rule;
        }
        public CFRule AddExpression(string formula, CellStyle style)
        {
            CFRule rule = new CFRule(CFType.Expression);
            rule.Formula1 = formula;
            rule.Style = style;
            AddRule(rule);
            return rule;
        }

        public CFRule AddContainsText(string value, CellStyle style)
        {
            CFType type = CFType.ContainsText;
            CFRule rule = GetRuleForOneValue(type, value, style, CFOperator.ContainsText, value);
            return rule;
        }

        public CFRule AddNotContainsText(string value, CellStyle style)
        {
            CFType type = CFType.NotContainsText;
            CFRule rule = GetRuleForOneValue(type, value, style, CFOperator.NotContains, value);
            return rule;
        }

        public CFRule AddBeginsWithText(string value, CellStyle style)
        {
            CFType type = CFType.BeginsWith;
            CFRule rule = GetRuleForOneValue(type, value, style, CFOperator.BeginsWith, value);
            return rule;
        }

        public CFRule AddEndsWithText(string value, CellStyle style)
        {
            CFType type = CFType.EndsWith;
            CFRule rule = GetRuleForOneValue(type, value, style, CFOperator.EndsWith, value);
            return rule;
        }

        public CFRule AddIsBlank(CellStyle style)
        {
            CFType type = CFType.ContainsBlanks;
            CFRule rule = new CFRule(type);
            rule.Formula1 = CFRule.CreateFormula(null, type, GetFirstCellStringValue());
            rule.Style = style;
            AddRule(rule);
            return rule;
        }

        public CFRule AddIsNotBlank(CellStyle style)
        {
            CFType type = CFType.NotContainsBlanks;
            CFRule rule = new CFRule(type);
            rule.Formula1 = CFRule.CreateFormula(null, type, GetFirstCellStringValue());
            rule.Style = style;
            AddRule(rule);
            return rule;
        }

        public CFRule AddContainsErrors(CellStyle style)
        {
            CFType type = CFType.ContainsErrors;
            CFRule rule = new CFRule(type);
            rule.Formula1 = CFRule.CreateFormula(null, type, GetFirstCellStringValue());
            rule.Style = style;
            AddRule(rule);
            return rule;
        }

        public CFRule AddNotContainsErrors(CellStyle style)
        {
            CFType type = CFType.NotContainsErrors;
            CFRule rule = new CFRule(type);
            rule.Formula1 = CFRule.CreateFormula(null, type, GetFirstCellStringValue());
            rule.Style = style;
            AddRule(rule);
            return rule;
        }

        public CFRule AddTimePeriod(TimePeriod timePeriodType, CellStyle style)
        {
            CFType type = CFType.TimePeriod;
            CFRule rule = new CFRule(type);
            rule.Formula1 = CFRule.CreateTimePeriodFormula(timePeriodType, GetFirstCellStringValue());
            rule.Style = style;
            rule.TimePeriod = timePeriodType;
            AddRule(rule);
            return rule;
        }

        #endregion

    }

    #region COLOR SCALE
    public class ColorScale
    {
        public List<CFVO> CFVOList { get; set; }
        public List<Color> Colors { get; set; }

        public ColorScale()
        {
            CFVOList = new List<CFVO>();
            Colors = new List<Color>();
        }

        public static ColorScale TwoColor(Color minimum, Color maximum)
        {
            ColorScale cs = new ColorScale();
            var cfvoMin = new CFVO
            {
                Type = CFVOType.Min
            };
            var cfvoMax = new CFVO
            {
                Type = CFVOType.Max
            };
            cs.CFVOList.Add(cfvoMin);
            cs.CFVOList.Add(cfvoMax);
            cs.Colors.Add(minimum);
            cs.Colors.Add(maximum);
            return cs;    
        }

        public static ColorScale ThreeColor(Color minimum, Color midpoint, Color maximum, int centerPercentage = 50)
        {
            ColorScale cs = new ColorScale
            {
                CFVOList = new List<CFVO>() { 
                    new CFVO { Type = CFVOType.Min },
                    new CFVO { Type = CFVOType.Percent, Value = centerPercentage.ToString() },
                    new CFVO { Type = CFVOType.Max }
                },
                Colors = new List<Color>() { minimum, midpoint, maximum }
            };
            return cs;
        }

        public static ColorScale DefaultTwoColor()
        {
            Color minimum = Color.Red, maximum = Color.Green;
            return TwoColor(minimum, maximum);
        }

        public static ColorScale DefaultThreeColor(int centerPercentage = 50)
        {
            Color minimum = Color.Red, maximum = Color.Green, midpoint = Color.Yellow;
            return ThreeColor(minimum, midpoint, maximum, centerPercentage);
        }
        
    }
    #endregion

    #region DATA BAR
    public class DataBar
    {
        public List<CFVO> CFVOList { get; set; }
        public Color Color { get; set; }

        public DataBar()
        {
            CFVOList = new List<CFVO>();
            Color = Color.Blue;
        }

        public static DataBar CreateDefaultDataBar(Color color)
        {
            DataBar db = new DataBar();
            db.Color = color;
            db.CFVOList = new List<CFVO>
            {
                new CFVO
                {
                    Type = CFVOType.Min
                },
                new CFVO
                {
                    Type = CFVOType.Max
                }
            };
            return db;
        }
    }
    #endregion

    #region ICON SET
    public class IconSet
    {
        public List<CFVO> CFVOList { get; set; }
        public IconSetType IconSetType { get; set; }
        public bool ShowValue { get; set; }

        public IconSet()
        {
            IconSetType = IconSetType.Item3Arrows;
            CFVOList = new List<CFVO>();
            ShowValue = true;
        }

        private static List<CFVO> GenerateCFVOs(IconSetType iconSetType) 
        {
            int numberOfIcons = GetNumberOfIcons(iconSetType);
            decimal step = 100.0m / numberOfIcons;
            List<CFVO> list = new List<CFVO>();
            for (int i = 0; i < numberOfIcons; i++)
            {
                int percent = (int)Math.Round(i * step);
                CFVO cfvo = new CFVO
                {
                    Type = CFVOType.Percent,
                    Value = percent.ToString()
                };
                list.Add(cfvo);
            }
            return list;
        }

        private static int GetNumberOfIcons(IconSetType iconSetType) 
        {
            IconSetType[] fourIcons = { 
                IconSetType.Item4Arrows,
                IconSetType.Item4ArrowsGray,
                IconSetType.Item4RedToBlack,
                IconSetType.Item4Rating,
                IconSetType.Item4TrafficLights
            };
            IconSetType[] fiveIcons = { 
                IconSetType.Item5Arrows,
                IconSetType.Item5ArrowsGray,
                IconSetType.Item5Rating,
                IconSetType.Item5Quarters
            };
            if (fourIcons.Contains(iconSetType))
                return 4;
            if (fiveIcons.Contains(iconSetType))
                return 5;
            return 3;
        }

        public static IconSet CreateDefaultIconSet(IconSetType iconSetType = IconSetType.Item3Arrows, bool showValue = true)
        {
            IconSet iconSet = new IconSet();
            iconSet.IconSetType = iconSetType;
            iconSet.CFVOList = GenerateCFVOs(iconSetType);
            iconSet.ShowValue = showValue;
            return iconSet;
        }
    }
    #endregion

    public class CFVO
    {
        public CFVO() { }
        public CFVO(CFVOType type, string value)
        {
            Type = type;
            Value = value;
        }
        public CFVOType Type { get; set; }
        public string Value { get; set; }
    }

    public enum IconSetType
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

    public enum TimePeriod
    {
        Today,
        Yesterday,
        Tomorrow,
        Last7Days,
        ThisMonth,
        LastMonth,
        NextMonth,
        ThisWeek,
        LastWeek,
        NextWeek
    }
     
    
}
