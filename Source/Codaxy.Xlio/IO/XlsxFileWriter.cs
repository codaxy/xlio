using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Codaxy.Xlio.Model.Oxml;
using Codaxy.Xlio.Model.Opc;
using System.Xml.Serialization;
using System.Globalization;
using System.Diagnostics;
using System.IO.Compression;

namespace Codaxy.Xlio.IO
{
    [Flags]
    public enum XlsxFileWriterOptions : int
    {
        None,
        AutoFit = 1
    }

    public partial class XlsxFileWriter : IDisposable
    {
        ZipArchive archive;
        Workbook workbook;
        public XlsxFileWriterOptions Options { get; set; }
        public XlsxFileWriter(Stream output)
        {
            this.archive = new ZipArchive(output, ZipArchiveMode.Create);
        }

        SharedStrings sharedStrings;
        Relationships rels, workbookRelationships;
        List<object> contentTypes;

        public void Write(Workbook workbook)
        {
            /* Write workbook.xml
             * Write workbook.xml.rels
             * Write sheets.xml
             * Register all content types
             */

            this.workbook = workbook;
            rels = new Relationships();
            contentTypes = new List<object>();

            WriteWorkbook();

            //Write package relationships            
            WriteRelationsips("_rels/.rels", rels);

            //Write [Content_Types].xml            
            CT_Types ctypes = new CT_Types { Items = contentTypes.ToArray() };
            WriteFile("[Content_Types].xml", ctypes, ContentTypesNs());
        }

        void OverrideContentType(String part, String type)
        {
            contentTypes.Add(new CT_Override { ContentType = type, PartName = "/" + part.TrimStart('/') });
        }

        void WriteWorkbook()
        {
            rels.Add(new CT_Relationship { Type = Relationships.Workbook, Target = "xl/workbook.xml" });
            OverrideContentType("/xl/workbook.xml", ContentTypes.Workbook);

            sharedStrings = new SharedStrings();
            ClearStyles();

            workbookRelationships = new Relationships();

            int sheetIndex = 1;

            contentTypes.Add(new CT_Default { Extension = "xml", ContentType = ContentTypes.Xml });
            contentTypes.Add(new CT_Default { Extension = "rels", ContentType = ContentTypes.Rels });

            var sheets = new List<CT_Sheet>();

            foreach (var sheet in workbook.Sheets)
            {
                var relativePath = String.Format("worksheets/sheet{0}.xml", sheetIndex);
                var sheetPath = "xl/" + relativePath;
                var rid = workbookRelationships.Add(new CT_Relationship
                {
                    Target = relativePath,
                    TargetMode = ST_TargetMode.Internal,
                    Type = Relationships.Sheet
                });
                OverrideContentType(sheetPath, ContentTypes.Sheet);
                var st = new CT_Sheet
                {
                    id = rid,
                    name = sheet.SheetName,
                    sheetId = (uint)sheetIndex
                };
                sheets.Add(st);
                WriteSheet(sheetPath, sheet);
                sheetIndex++;
            }

            var wb = new CT_Workbook()
            {
                sheets = sheets.ToArray(),
                bookViews = new[] { new CT_BookView() }
            };
            
            var definedNames = new List<CT_DefinedName>();
            foreach (var definedName in workbook.DefinedNames)
            {
                var dn = new CT_DefinedName
                {
                    name = definedName.Name,
                    Value = definedName.Value
                };
                definedNames.Add(dn);
            }

            for (var i = 0; i < workbook.Sheets.Count; i++)
            {
                foreach (var definedName in workbook.Sheets[i].DefinedNames)
                {
                    var dn = new CT_DefinedName
                    {
                        name = definedName.Name,
                        Value = definedName.Value,
                        localSheetId = (uint)i,
                        localSheetIdSpecified = true
                    };
                    definedNames.Add(dn);
                }
            }

            if (definedNames.Count > 0)
                wb.definedNames = definedNames.ToArray();
            

            WriteFile("xl/workbook.xml", wb, SpreadsheetNs(true));

            if (sharedStrings.Count > 0)
                WriteSharedStrings();

            if (!styles.Empty)
            {
                OverrideContentType("xl/styles.xml", ContentTypes.Styles);
                workbookRelationships.Add(new CT_Relationship { Target = "styles.xml", Type = Relationships.Styles });
                WriteStyles("xl/styles.xml");
            }

            WriteRelationsips("xl/_rels/workbook.xml.rels", workbookRelationships);
        }

        private void WriteStyles(string path)
        {
            CT_Stylesheet stylesheet = new CT_Stylesheet
            {
                cellXfs = new CT_CellXfs { xf = styles.ToArray() },
                cellStyles = new CT_CellStyles { cellStyle = new[] { new CT_CellStyle { name = "Normal", xfId = 0, builtinId = 0, builtinIdSpecified = true } } },
                cellStyleXfs = new CT_CellStyleXfs { xf = new[] { new CT_Xf() } },
                borders = new CT_Borders { border = borders.ToArray() },
                fonts = new CT_Fonts { font = fonts.ToArray() },
                dxfs = new CT_Dxfs { dxf = dxfStyles.ToArray() },
                colors = new CT_Colors(),
                fills = new CT_Fills { fill = fills.ToArray() },
                numFmts = new CT_NumFmts { numFmt = numFormats.ToArray() },
                tableStyles = new CT_TableStyles()
            };

            WriteFile(path, stylesheet, SpreadsheetNs(false));
        }

        private void WriteRelationsips(String path, Relationships rels)
        {
            CT_Relationships relo = new CT_Relationships() { Relationship = rels.ToArray() };
            WriteFile(path, relo, RelsNs());
        }

        private void WriteSharedStrings()
        {
            var rest = new List<CT_Rst>();
            foreach (var s in sharedStrings)
            {
                rest.Add(new CT_Rst
                {
                    t = s,
                    Space = s.StartsWith(" ") || s.EndsWith(" ") ? "preserve" : null
                });
            }
            CT_Sst sst = new CT_Sst() { si = rest.ToArray() };
            var relativePath = "sharedStrings.xml";
            var path = "xl/" + relativePath;
            WriteFile(path, sst, SpreadsheetNs(false));
            workbookRelationships.Add(new CT_Relationship { Type = Relationships.SharedStrings, Target = relativePath });
            OverrideContentType(path, ContentTypes.SharedStrings);
        }

        CT_MergeCell ConvertRangeToMergedCell(Range range)
        {
            return new CT_MergeCell { @ref = range.ToString() };
        }

        private void WriteSheet(string sheetPath, Sheet sheet)
        {
            var sharedFormulas = new Dictionary<SharedFormula, int>();
            var mergedCells = new Mapper<Range, CT_MergeCell>(ConvertRangeToMergedCell);

            bool[] autoFit = new bool[256];
            int[] width = new int[256];
            bool calcAutoFit = (Options & XlsxFileWriterOptions.AutoFit) != 0;


            var sheetFormat = new Lazy<CT_SheetFormatPr>();
            if (sheet.DefaultRowHeight.HasValue)
            {
                sheetFormat.Value.defaultRowHeight = sheet.DefaultRowHeight.Value;
                sheetFormat.Value.customHeight = true;
            }

            if (sheet.DefaultColumnWidth.HasValue)
            {
                sheetFormat.Value.defaultColWidth = sheet.DefaultColumnWidth.Value;
                sheetFormat.Value.defaultColWidthSpecified = true;
            }

            var sheetView = new CT_SheetView();

            if (sheet.ActiveCell != null)
                sheetView.selection = new[] { new CT_Selection { activeCell = sheet.ActiveCell.ToString(), sqref = new[] { sheet.ActiveCell.ToString() } } };

            sheetView.showGridLines = sheet.ShowGridLines;


            var rows = new List<CT_Row>();
            foreach (var row in sheet.Data)
            {
                var cells = new List<CT_Cell>();

                foreach (var cd in row.Value)
                {
                    bool skipData = false;
                    var c = new Cell { Col = cd.Key, Row = row.Key };
                    if (cd.Value.IsMerged)
                        if (cd.Value.MergedRange.Cell1 == c)
                            mergedCells.GetIndex(cd.Value.MergedRange);
                        else
                            skipData = true;

                    ST_CellType ct;
                    var data = cd.Value;

                    var cell = new CT_Cell
                    {
                        r = c.ToString()
                    };

                    if (!skipData)
                    {
                        if (data.Formula != null)
                        {
                            if (data.Formula.StartsWith("{"))
                            {
                                cell.f = new CT_CellFormula
                                {
                                    t = ST_CellFormulaType.array,
                                    Value = data.Formula.Substring(1, data.Formula.Length - 2)
                                };
                            }
                            else
                                cell.f = new CT_CellFormula { Value = data.Formula };
                        }
                        else if (data.SharedFormula != null)
                        {
                            int si;
                            if (!sharedFormulas.TryGetValue(data.SharedFormula, out si))
                            {
                                si = sharedFormulas.Count; //zero based
                                sharedFormulas.Add(data.SharedFormula, si);
                                cell.f = new CT_CellFormula
                                {
                                    t = ST_CellFormulaType.shared,
                                    Value = data.SharedFormula.Formula,
                                    @ref = data.SharedFormula.Range.ToString(),
                                    si = (uint)si,
                                    siSpecified = true,
                                };
                            }
                            else
                                cell.f = new CT_CellFormula
                                {
                                    si = (uint)si,
                                    siSpecified = true,
                                    t = ST_CellFormulaType.shared
                                };
                        }
                        else
                        {
                            String sv, format;
                            var v = WriteCellValue(data, out ct, out sv, out format);
                            if (format != null && (data.style == null || data.style.format == null))
                                data.Style.Format = format;
                            cell.v = v;
                            cell.t = ct;
                            if (calcAutoFit && sv != null && !cd.Value.IsMerged && cd.Key < 256)
                            {
                                autoFit[cd.Key] = true;
                                if (sv.Length > width[cd.Key])
                                    width[cd.Key] = sv.Length;
                            }
                        }
                    }

                    if (data.style != null)
                        cell.s = RegisterStyle(data.style);

                    cells.Add(cell);
                }

                var r = new CT_Row()
                {
                    c = cells.ToArray(),
                    r = (uint)row.Key + 1,
                    rSpecified = true
                };

                if (row.Value.Height.HasValue)
                {
                    r.ht = row.Value.Height.Value;
                    r.htSpecified = true;
                    r.customHeight = true;
                }

                r.hidden = row.Value.Hidden;
                r.collapsed = row.Value.Collapsed;
                r.ph = row.Value.Phonetic;

                if (row.Value.style != null)
                {
                    r.s = RegisterStyle(row.Value.style);
                    r.customFormat = r.s > 0;
                }

                rows.Add(r);
            }

            var cols = new List<CT_Col>();
            if (calcAutoFit)
            {
                for (int i = 0; i < width.Length; i++)
                    if (width[i] > 0)
                    {
                        if (!sheet.Columns[i].Width.HasValue)
                        {
                            sheet.Columns[i].Width = width[i];
                            sheet.Columns[i].BestFit = true;
                        }

                    }
            }

            if (sheet.Columns.data.Count > 0)
            {
                foreach (var node in sheet.Columns.data)
                {
                    var col = new CT_Col
                    {
                        min = (uint)(node.FirstColumn + 1),
                        max = (uint)(node.LastColumn + 1),
                        bestFit = node.Data.BestFit,
                        hidden = node.Data.Hidden,
                        outlineLevel = node.Data.OutlineLevel,
                        phonetic = node.Data.Phonetic,
                        width = node.Data.Width ?? 10,
                        widthSpecified = true,
                        customWidth = node.Data.Width.HasValue
                    };

                    if (node.Data.style != null)
                        col.style = RegisterStyle(node.Data.style);

                    cols.Add(col);
                }
            }

            var ws = new CT_Worksheet
            {
                sheetData = rows.ToArray(),
                cols = cols.Count > 0 ? cols.ToArray() : null,
                sheetFormatPr = sheetFormat.Data,
                sheetViews = new CT_SheetViews { sheetView = new[] { sheetView } },
                pageSetup = new CT_PageSetup() { scale = (uint)sheet.Page.Scale }
            };

            switch (sheet.Page.Orientation)
            {
                case PageOrientation.Landscape:
                    ws.pageSetup.orientation = ST_Orientation.landscape;
                    break;
                case PageOrientation.Portrait:
                    ws.pageSetup.orientation = ST_Orientation.portrait;
                    break;
            }

            if (sheet.Page.Margins != null)
            {
                ws.pageMargins = new CT_PageMargins
                {
                    bottom = sheet.Page.Margins.Bottom,
                    footer = sheet.Page.Margins.Footer,
                    header = sheet.Page.Margins.Header,
                    left = sheet.Page.Margins.Left,
                    right = sheet.Page.Margins.Right,
                    top = sheet.Page.Margins.Top
                };
            }

            if (sheet.ColumnBreaks != null && sheet.ColumnBreaks.Count > 0)
            {
                ws.colBreaks = new CT_PageBreak
                {
                    count = (uint)sheet.ColumnBreaks.Count,
                    manualBreakCount = (uint)sheet.ColumnBreaks.Count,
                    brk = sheet.ColumnBreaks.Select(id => new CT_Break
                    {
                        id = (uint)id + 1,
                        man = true,
                        //max = (uint)sheet.Data.LastRow + 1
                    }).ToArray()
                };
            }

            if (sheet.RowBreaks != null && sheet.RowBreaks.Count > 0)
            {
                ws.rowBreaks = new CT_PageBreak
                {
                    count = (uint)sheet.RowBreaks.Count,
                    manualBreakCount = (uint)sheet.RowBreaks.Count,
                    brk = sheet.RowBreaks.Select(id => new CT_Break
                    {
                        id = (uint)id + 1,
                        man = true,
                        //max = 1048575U
                    }).ToArray()
                };
            }

            if (!mergedCells.Empty)
                ws.mergeCells = new CT_MergeCells { mergeCell = mergedCells.ToArray() };

            //conditional formatting
            CT_ConditionalFormatting[] cf;
            WriteConditionalFormatting(sheet, out cf);
            ws.conditionalFormatting = cf;

            WriteFile(sheetPath, ws, SpreadsheetNs(false));
        }

        #region ConditionalFormatting

        public void WriteConditionalFormatting(Sheet sheet, out CT_ConditionalFormatting[] cf)
        {
            List<CT_ConditionalFormatting> cflist = new List<CT_ConditionalFormatting>();
            foreach (var cfrCollection in sheet.ConditionalFormatting)
                if (cfrCollection.Ranges != null && cfrCollection.Ranges.Count > 0)
                {
                    List<CT_CfRule> ct_cfrule_list = new List<CT_CfRule>();
                    foreach (var ruleobj in cfrCollection)
                    {
                        CT_CfRule cfRule = null;
                        WriteConditionalFormattingRule(ruleobj, cfrCollection.GetFirstCellStringValue(), out cfRule);
                        ct_cfrule_list.Add(cfRule);
                    }
                    var tempCf = new CT_ConditionalFormatting
                    {
                        cfRule = ct_cfrule_list.ToArray(),
                        sqref = cfrCollection.Ranges.ToArray()
                    };
                    cflist.Add(tempCf);
                }
            cf = cflist.ToArray();
        }

        public void WriteConditionalFormattingRule(CFRule rule, string firstCell, out CT_CfRule cfRule)
        {
            ST_CfType st_cftype = (ST_CfType) rule.Type;
            CT_ColorScale ct_colorscale = null;
            CT_DataBar ct_databar = null;
            CT_IconSet ct_iconset = null;

            if (rule.Type == CFType.ColorScale)
                WriteConditionalFormattingColorScale(rule.ColorScale, out ct_colorscale);

            if (rule.Type == CFType.DataBar)
                WriteConditionalFormattingDataBar(rule.DataBar, out ct_databar);

            if (rule.Type == CFType.IconSet)
                WriteConditionalFormattingIconSet(rule.IconSet, out ct_iconset);

            if (CFRule.RequiresDxf(rule.Type))
            {
                List<string> formulas = new List<string>();

                if (!string.IsNullOrEmpty(rule.Formula1))
                    formulas.Add(rule.Formula1);

                if (!string.IsNullOrEmpty(rule.Formula2))
                    formulas.Add(rule.Formula2);

                if (!string.IsNullOrEmpty(rule.Formula3))
                    formulas.Add(rule.Formula3);

                cfRule = new CT_CfRule
                {
                    typeSpecified = true,
                    type = st_cftype,
                    priority = rule.Priority,
                    formula = formulas.ToArray(),
                    operatorSpecified = rule.Operator.HasValue,
                    @operator = rule.Operator.HasValue ? (ST_ConditionalFormattingOperator)rule.Operator.Value : ST_ConditionalFormattingOperator.beginsWith,
                    text = rule.Text,
                };
                cfRule = SetDxfStyle(cfRule, rule.Style);

                switch (rule.Type)
                {
                    case CFType.Top10:
                        cfRule.percent = rule.IsPercent;
                        cfRule.bottom = rule.IsBottom;
                        cfRule.rankSpecified = rule.Rank.HasValue;
                        cfRule.rank = (uint)rule.Rank.Value;
                        break;

                    case CFType.AboveAverage:
                        cfRule.aboveAverage = rule.IsAboveAverage;
                        cfRule.equalAverage = rule.IsEqualAverage;
                        cfRule.stdDevSpecified = rule.IsStdDev;
                        cfRule.stdDev = rule.StdDev;
                        break;

                    case CFType.TimePeriod:
                        cfRule.timePeriodSpecified = true;
                        cfRule.timePeriod = (ST_TimePeriod)rule.TimePeriod;
                        break;
                }
                return;
            }
            
            cfRule = new CT_CfRule
            {
                typeSpecified = true,
                type = st_cftype,
                colorScale = ct_colorscale,
                dataBar = ct_databar,
                iconSet = ct_iconset,
                priority = rule.Priority
            };

        }

        public CT_CfRule SetDxfStyle(CT_CfRule ct_cfrule, CellStyle style)
        {
            if (style != null)
            {
                ct_cfrule.dxfId = RegisterDxfStyle(style);
                ct_cfrule.dxfIdSpecified = true;
            }
            return ct_cfrule;
        }

        public void WriteConditionalFormattingColorScale(ColorScale colorScale, out CT_ColorScale cfColorScale)
        {
            if (colorScale == null)
            {
                cfColorScale = null;
                return;
            }
            List<CT_Cfvo> ct_cfvo_list = new List<CT_Cfvo>();
            foreach (var ct_cfvo in colorScale.CFVOList)
            {
                ct_cfvo_list.Add(new CT_Cfvo
                {
                    val = ct_cfvo.Value,
                    type = (ST_CfvoType)ct_cfvo.Type //1-1
                });
            }

            List<CT_Color1> ct_color1_list = new List<CT_Color1>();
            foreach (var color in colorScale.Colors)
            {
                ct_color1_list.Add(ConvertColor(color));
            }

            cfColorScale = new CT_ColorScale
            {
                cfvo = ct_cfvo_list.ToArray(),
                color = ct_color1_list.ToArray()
            };
        }

        public void WriteConditionalFormattingDataBar(DataBar dataBar, out CT_DataBar cfDataBar)
        {
            if (dataBar == null)
            {
                cfDataBar = null;
                return;
            }
            List<CT_Cfvo> ct_cfvo_list = new List<CT_Cfvo>();
            foreach (var ct_cfvo in dataBar.CFVOList)
            {
                ct_cfvo_list.Add(new CT_Cfvo
                {
                    val = ct_cfvo.Value,
                    type = (ST_CfvoType)ct_cfvo.Type //1-1
                });
            }
            cfDataBar = new CT_DataBar
            {
                cfvo = ct_cfvo_list.ToArray(),
                color = ConvertColor(dataBar.Color)
            };
        }

        public void WriteConditionalFormattingIconSet(IconSet iconSet, out CT_IconSet cfIconSet)
        {
            if (iconSet == null)
            {
                cfIconSet = null;
                return;
            }
            List<CT_Cfvo> ct_cfvo_list = new List<CT_Cfvo>();
            foreach (var ct_cfvo in iconSet.CFVOList)
            {
                ct_cfvo_list.Add(new CT_Cfvo
                {
                    val = ct_cfvo.Value,
                    type = (ST_CfvoType)ct_cfvo.Type //1-1
                });
            }

            cfIconSet = new CT_IconSet
            {
                cfvo = ct_cfvo_list.ToArray(),
                iconSet = (ST_IconSetType)iconSet.IconSetType,//1-1
                showValue = iconSet.ShowValue
            };
        }

        #endregion

        private string WriteCellValue(CellData data, out ST_CellType ct, out String value, out String format)
        {
            format = null;

            if (data.Value == null)
            {
                ct = ST_CellType.n;
                value = null;
                return null;
            }
            var type = data.Value.GetType();
            if (TypeInfo.IsNumericType(type))
            {
                ct = ST_CellType.n;
                return value = String.Format(CultureInfo.InvariantCulture, "{0}", data.Value);
            }
            if (type == typeof(DateTime))
            {
                var dt = (DateTime)data.Value;
                ct = ST_CellType.n;
                format = "mm-dd-yy";
                return value = String.Format(CultureInfo.InvariantCulture, "{0}", XlioUtil.ToExcelDateTime(dt));
            }
            if (type == typeof(string))
            {
                ct = ST_CellType.s;
                value = (String)data.Value;
                return sharedStrings.Get(value).ToString();
            }
            if (type == typeof(bool))
            {
                ct = ST_CellType.b;
                return value = (bool)data.Value == true ? "1" : "0";
            }

            ct = ST_CellType.inlineStr;
            return value = String.Format(CultureInfo.InvariantCulture, "{0}", data.Value);
        }

        XmlSerializerNamespaces SpreadsheetNs(bool relsNs)
        {
            XmlSerializerNamespaces res = new XmlSerializerNamespaces();
            res.Add("", XmlNs.SpreadsheetML);
            if (relsNs)
                res.Add("r", XmlNs.Relationships);
            return res;
        }

        XmlSerializerNamespaces RelsNs()
        {
            XmlSerializerNamespaces res = new XmlSerializerNamespaces();
            res.Add("", XmlNs.Package.Relationships);
            return res;
        }

        XmlSerializerNamespaces ContentTypesNs()
        {
            XmlSerializerNamespaces res = new XmlSerializerNamespaces();
            res.Add("", XmlNs.Package.ContentTypes);
            return res;
        }

        private void WriteFile<T>(string filePath, T data, XmlSerializerNamespaces ns)
        {
            var entry = archive.CreateEntry(filePath);
            var xs = new XmlSerializer(typeof(T));
            using (var output = entry.Open())
            {
                var writer = new StreamWriter(output, Encoding.UTF8);
                xs.Serialize(writer, data, ns);
            }
        }

        public void Dispose()
        {
            this.archive.Dispose();
            this.archive = null;
        }

        struct Lazy<T> where T : class, new()
        {
            T v;

            public T Value
            {
                get
                {
                    return v ?? (v = new T());
                }
            }

            public T Data
            {
                get { return v; }
            }
        }


    }
}
