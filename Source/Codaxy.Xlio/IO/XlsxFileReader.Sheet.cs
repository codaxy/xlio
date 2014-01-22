using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio.Model.Oxml;
using System.Diagnostics;
using System.Globalization;

namespace Codaxy.Xlio.IO
{
    public partial class XlsxFileReader
    {
        private void ReadSheet(string sheetPath, Sheet sheet)
        {
            Dictionary<uint, SharedFormula> sharedFormulas = new Dictionary<uint, SharedFormula>();
            try
            {
                var xml = ReadFile<CT_Worksheet>(sheetPath);

                if (xml.sheetFormatPr != null)
                {
                    if (xml.sheetFormatPr.customHeight)
                        sheet.DefaultRowHeight = xml.sheetFormatPr.defaultRowHeight;

                    if (xml.sheetFormatPr.defaultColWidthSpecified)
                        sheet.DefaultColumnWidth = xml.sheetFormatPr.defaultColWidth;
                }

                if (xml.sheetViews!=null && xml.sheetViews.sheetView!=null)
                    foreach (var sheetView in xml.sheetViews.sheetView)
                        if (sheetView.workbookViewId == 0) //default view
                        {
                            sheet.ShowGridLines = sheetView.showGridLines;

                            if (sheetView.selection != null && sheetView.selection.Length == 1)
                            {
                                if (sheetView.selection[0].activeCell != null)
                                    sheet.ActiveCell = Cell.Parse(sheetView.selection[0].activeCell);
                            }
                        }

                if (xml.cols != null)
                {
                    foreach (var col in xml.cols)
                    {
                        sheet.Columns.AppendRange((int)col.min - 1, (int)col.max - 1, new SheetColumn
                        {
                            BestFit = col.bestFit,
                            Width = col.customWidth && col.widthSpecified ? col.width : (double?)null,
                            style = GetStyle(col.style),
                            Phonetic = col.phonetic,
                            OutlineLevel = col.outlineLevel,
                            Hidden = col.hidden
                        });
                    }
                }

                if (xml.colBreaks != null)
                    foreach (var b in xml.colBreaks.brk)
                        if (b.man)
                            sheet.ColumnBreaks.Add((int)b.id - 1);

                if (xml.mergeCells!=null && xml.mergeCells.mergeCell != null)
                    foreach (var mc in xml.mergeCells.mergeCell)
                    {
                        var range = Range.Parse(mc.@ref);
                        sheet[range].Merge();
                    }
                
                if (xml.sheetData!=null)
                    foreach (var rd in xml.sheetData)
                    {
                        var row = sheet[(int)rd.r - 1];

                        if (rd.customFormat)
                            row.Style = GetStyle(rd.s);

                        if (rd.customHeight && rd.htSpecified)
                            row.Height = rd.ht;

                        row.Phonetic = rd.ph;
                        row.Collapsed = rd.collapsed;
                        row.Hidden = rd.hidden;

                        if (rd.c != null)
                            foreach (var c in rd.c)
                            {
                                var cell = Cell.Parse(c.r);
                                var data = row[cell.Col];

                                data.Style = GetStyle(c.s);

                                if (c.v != null)
                                {
                                    object value = null;
                                    switch (c.t)
                                    {
                                        case ST_CellType.n:
                                            var n = Convert.ToDouble(c.v, CultureInfo.InvariantCulture);
                                            value = data.style != null && NumberFormat.IsDateTimeFormat(data.style.format) ? value = XlioUtil.ToDateTime(n) : n;
                                            break;
                                        case ST_CellType.inlineStr:
                                            value = c.v;
                                            break;
                                        case ST_CellType.s:
                                            value = sharedStrings[Convert.ToInt32(c.v, CultureInfo.InvariantCulture)];
                                            break;
                                        case ST_CellType.b:
                                            value = c.v != null && c.v.Length > 0 && (c.v[0] == 'T' || c.v[0] == 't' || c.v[0] == '1');
                                            break;
                                    }
                                    row[cell.Col].Value = value;
                                }

                                if (c.f != null)
                                {
                                    switch (c.f.t)
                                    {
                                        case ST_CellFormulaType.normal:
                                            row[cell.Col].Formula = c.f.Value;
                                            break;
                                        case ST_CellFormulaType.shared:
                                            if (c.f.siSpecified)
                                            {
                                                SharedFormula sharedFormula = null;
                                                if (c.f.@ref != null && c.f.Value != null) //shared formula definition
                                                {
                                                    var range = Range.Parse(c.f.@ref);
                                                    var origin = cell;
                                                    var formula = c.f.Value;
                                                    sharedFormula = new SharedFormula(formula, range, origin);
                                                    sharedFormulas[c.f.si] = sharedFormula;
                                                }
                                                else
                                                    sharedFormulas.TryGetValue(c.f.si, out sharedFormula);

                                                row[cell.Col].SharedFormula = sharedFormula;
                                            }
                                            break;
                                        case ST_CellFormulaType.array:
                                            row[cell.Col].Formula = "{" + c.f.Value + "}";
                                            break;
                                    }
                                }

                            }

                        if (xml.rowBreaks != null)
                            foreach (var brk in xml.rowBreaks.brk)
                                if (brk.man)
                                    sheet.RowBreaks.Add((int)brk.id - 1);

                        if (xml.pageSetup != null)
                        {
                            switch (xml.pageSetup.orientation)
                            {
                                case ST_Orientation.portrait:
                                    sheet.Page.Orientation = PageOrientation.Portrait;
                                    break;
                                case ST_Orientation.landscape:
                                    sheet.Page.Orientation = PageOrientation.Landscape;
                                    break;
                            }

                            if (xml.pageSetup.scale > 0)
                                sheet.Page.Scale = (int)xml.pageSetup.scale;
                        }

                        if (xml.pageMargins != null)
                        {
                            if (xml.pageMargins != null)
                                sheet.Page.Margins = new PageMargins
                                {
                                    Bottom = xml.pageMargins.bottom,
                                    Footer = xml.pageMargins.footer,
                                    Header = xml.pageMargins.header,
                                    Left = xml.pageMargins.left,
                                    Right = xml.pageMargins.right,
                                    Top = xml.pageMargins.top
                                };
                        }
                    }

                //conditional formatting
                if (xml.conditionalFormatting != null)
                foreach (CT_ConditionalFormatting conditionalFormatting in xml.conditionalFormatting)
                {
                    var rules = new List<ConditionalFormattingRule>();
                    if (conditionalFormatting.cfRule !=null)
                    foreach (var cfRule in conditionalFormatting.cfRule)
                    {
                        var rule = GetConditionalFormattingRule(cfRule);
                        rules.Add(rule);
                    }
                    ConditionalFormatting cf = new ConditionalFormatting
                    {
                        Ranges = conditionalFormatting.sqref.ToList(),
                        Rules = rules
                    };
                    sheet.ConditionalFormatting.Add(cf);
                }
                 
            }
            catch (Exception ex)
            {                
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
                throw;
            }
        }

        private ConditionalFormattingRule GetConditionalFormattingRule(CT_CfRule ct_cfrule)
        {
            if (ct_cfrule.type == ST_CfType.iconSet)
            {
                return new IconSet
                {
                    IconSetType = (ConditionalFormattingIconSetType)ct_cfrule.iconSet.iconSet,
                    Priority = ct_cfrule.priority,
                    ShowValue = ct_cfrule.iconSet.showValue,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    CFVOList = ConvertCFVOList(ct_cfrule.iconSet.cfvo)
                };
            }
            if (ct_cfrule.type == ST_CfType.dataBar)
            {
                return new DataBar
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    CFVOList = ConvertCFVOList(ct_cfrule.dataBar.cfvo),
                    Color = ConvertColor(ct_cfrule.dataBar.color)
                };
            }
            if (ct_cfrule.type == ST_CfType.colorScale)
            {
                var colors = new List<Color>();
                foreach (CT_Color1 ct_color1 in ct_cfrule.colorScale.color)
                {
                    Color color = ConvertColor(ct_color1);
                    colors.Add(color);
                }
                return new ColorScale
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    CFVOList = ConvertCFVOList(ct_cfrule.colorScale.cfvo),
                    Colors = colors
                };
            }
            if (ct_cfrule.type == ST_CfType.cellIs)
            {
                string for1 = null, for2 = null, for3 = null;
                if (ct_cfrule.formula.Length > 0)
                    for1 = ct_cfrule.formula[0];
                if (ct_cfrule.formula.Length > 1)
                    for2 = ct_cfrule.formula[1];
                if (ct_cfrule.formula.Length > 2)
                    for3 = ct_cfrule.formula[2];
                return new ConditionalFormattingCondition
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    Formula1 = for1,
                    Formula2 = for2,
                    Formula3 = for3,
                    dxfId = ct_cfrule.dxfId,
                    Operator = (ConditionalFormattingOperator)ct_cfrule.@operator,
                    Text = ct_cfrule.text,

                };

            }
            //<cfRule type="top10" dxfId="1" priority="1" percent="1" bottom="1" rank="50"/>
            if (ct_cfrule.type == ST_CfType.top10)
            {
                return new ConditionalFormattingCondition
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    dxfId = ct_cfrule.dxfId,
                    IsPercent = ct_cfrule.percent,
                    IsBottom = ct_cfrule.bottom,
                    Rank = ct_cfrule.rank
                };
            }
            if (ct_cfrule.type == ST_CfType.aboveAverage)
            {
                return new ConditionalFormattingCondition
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    dxfId = ct_cfrule.dxfId,
                    IsAboveAverage = ct_cfrule.aboveAverage,
                    IsEqualAverage = ct_cfrule.equalAverage,
                    IsStdDev = ct_cfrule.stdDevSpecified,
                    StdDev = ct_cfrule.stdDev
                };
            }
            if (ct_cfrule.type == ST_CfType.uniqueValues || ct_cfrule.type == ST_CfType.duplicateValues)
            {
                return new ConditionalFormattingCondition
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    dxfId = ct_cfrule.dxfId
                };
            }
            if (ct_cfrule.type == ST_CfType.expression)
            {
                return new ConditionalFormattingCondition
                {
                    Priority = ct_cfrule.priority,
                    Type = (ConditionalFormattingType)ct_cfrule.type,
                    dxfId = ct_cfrule.dxfId,
                    Formula1 = ct_cfrule.formula[0]
                };
            }
            return null;
        }

        private List<ConditionalFormattingValueObject> ConvertCFVOList(CT_Cfvo[] cfvoArr)
        {
            var cfvoList = new List<ConditionalFormattingValueObject>();
            foreach (var cfvo in cfvoArr)
            {
                ConditionalFormattingValueObject CFVO = new ConditionalFormattingValueObject
                {
                    Type = (CFVOType)cfvo.type,
                    Value = cfvo.val
                };
                cfvoList.Add(CFVO);
            }
            return cfvoList;
        }
        
    }
}
