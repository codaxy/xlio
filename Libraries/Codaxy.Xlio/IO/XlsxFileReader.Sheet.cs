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
                var sheetData = ReadFile<CT_Worksheet>(sheetPath);

                if (sheetData.mergeCells!=null && sheetData.mergeCells.mergeCell != null)
                    foreach (var mc in sheetData.mergeCells.mergeCell)
                    {
                        var range = Range.Parse(mc.@ref);
                        sheet[range].Merge();
                    }
                
                if (sheetData.sheetData!=null)
                foreach (var rd in sheetData.sheetData)
                {
                    var row = sheet[(int)rd.r - 1];
                    if (rd.c!=null)
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
                }
                
            }
            catch (Exception ex)
            {                
                Debug.WriteLine(ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }
        
    }
}
