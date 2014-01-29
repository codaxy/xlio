using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class CFRule
    {
        public CFRule(CFType type, int priority = 1)
        {
            Priority = priority;
            Type = type;
            Operator = null;
            Formula1 = null;
            Formula2 = null;
            Formula3 = null;
            IsBottom = false;
            Rank = 10;
            IsPercent = false;
            StdDev = 0;
            IsAboveAverage = true;
            IsEqualAverage = false;
            IsStdDev = false;            
            Text = null;
            TimePeriod = TimePeriod.Today;
        }

        public CFRule(CFType type, CellStyle style, CFOperator @operator = CFOperator.LessThan, 
            string formula1 = "", string formula2 = "", string formula3 = "") 
        {
            Type = type;
            Style = style;
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
            
            Text = null;
            TimePeriod = TimePeriod.Today;
        }
        
        public int? Rank { get; set; }
        public bool IsBottom { get; set; }
        public bool IsPercent { get; set; }
        public bool IsAboveAverage { get; set; }
        public bool IsEqualAverage { get; set; }
        public bool IsStdDev { get; set; }
        
        public int StdDev { get; set; }
        public string Text { get; set; }
        public TimePeriod TimePeriod { get; set; }
        public CellStyle Style { get; set; }

        internal static bool RequiresDxf(CFType type)
        {
            switch (type)
            {
                case CFType.IconSet:                    
                case CFType.ColorScale:
                case CFType.DataBar:
                    return false;

                default:
                    return true;
            }
        }

        public string Formula1
        {
            get { return _formula1; }
            set { _formula1 = RemoveEqualChar(value); }
        }
        public string Formula2
        {
            get { return _formula2; }
            set { _formula2 = RemoveEqualChar(value); }
        }
        public string Formula3
        {
            get { return _formula3; }
            set { _formula3 = RemoveEqualChar(value); }
        }

        private string _formula1;
        private string _formula2;
        private string _formula3;

        private string RemoveEqualChar(string formula)
        {
            if (formula == null) return null;
            if (formula.StartsWith("="))
                return formula.Substring(1);
            return formula;
        }

        public ColorScale ColorScale { get; set; }

        public DataBar DataBar { get; set; }

        public IconSet IconSet { get; set; }

        public CFOperator? Operator { get; set; }
         
        public CFType Type { get; set; }

        public int Priority { get; set; }

        #region Formulas

        private static string containsText = "NOT(ISERROR(SEARCH(\"{0}\",{1})))";
        private static string notContainsText = "ISERROR(SEARCH(\"{0}\",{1}))";
        private static string beginsWith = "LEFT({1},LEN(\"{0}\"))=\"{0}\"";
        private static string endsWith = "RIGHT({1},LEN(\"{0}\"))=\"{0}\"";
        private static string containsBlanks = "LEN(TRIM({0}))=0";
        private static string notContainsBlanks = "LEN(TRIM({0}))>0";
        private static string containsErrors = "ISERROR({0})";
        private static string notContainsErrors = "NOT(ISERROR({0}))";

        private static string yesterday = "FLOOR({0},1)=TODAY()-1";
        private static string today = "FLOOR({0},1)=TODAY()";
        private static string tomorrow = "FLOOR({0},1)=TODAY()+1";
        private static string lastWeek = "AND(TODAY()-ROUNDDOWN({0},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({0},0)<(WEEKDAY(TODAY())+7))";
        private static string last7Days = "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())";
        private static string thisWeek = "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))";
        private static string nextWeek = "AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))";
        private static string lastMonth = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))";
        private static string thisMonth = "AND(MONTH({0})=MONTH(TODAY()),YEAR({0})=YEAR(TODAY()))";
        private static string nextMonth = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)),YEAR({0})=YEAR(EDATE(TODAY(),0+1)))";

        #endregion

        public static string CreateTimePeriodFormula(TimePeriod type, string firstCell)
        {
            switch (type)
            {
                case TimePeriod.Yesterday:
                    return string.Format(yesterday, firstCell);

                case TimePeriod.Today:
                    return string.Format(today, firstCell);

                case TimePeriod.Tomorrow:
                    return string.Format(tomorrow, firstCell);

                case TimePeriod.Last7Days:
                    return string.Format(last7Days, firstCell);

                case TimePeriod.LastWeek:
                    return string.Format(lastWeek, firstCell);

                case TimePeriod.ThisWeek:
                    return string.Format(thisWeek, firstCell);

                case TimePeriod.NextWeek:
                    return string.Format(nextWeek, firstCell);

                case TimePeriod.LastMonth:
                    return string.Format(lastMonth, firstCell);

                case TimePeriod.ThisMonth:
                    return string.Format(thisMonth, firstCell);

                case TimePeriod.NextMonth:
                    return string.Format(nextMonth, firstCell);

                default:
                    return firstCell;
            }
        }

        public static string CreateFormula(string value, CFType type, string firstCell)
        {
            switch (type) 
            {
                case CFType.ContainsText:
                    return string.Format(containsText, value, firstCell);

                case CFType.NotContainsText:
                    return string.Format(notContainsText, value, firstCell);

                case CFType.BeginsWith:
                    return string.Format(beginsWith, value, firstCell);

                case CFType.EndsWith:
                    return string.Format(endsWith, value, firstCell);

                case CFType.ContainsBlanks:
                    return string.Format(containsBlanks, firstCell);

                case CFType.NotContainsBlanks:
                    return string.Format(notContainsBlanks, firstCell);

                case CFType.ContainsErrors:
                    return string.Format(containsErrors, firstCell);

                case CFType.NotContainsErrors:
                    return string.Format(notContainsErrors, firstCell);

                default:
                    return value;
            }
        }
    }

    public enum CFType
    {
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

    public enum CFOperator
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
