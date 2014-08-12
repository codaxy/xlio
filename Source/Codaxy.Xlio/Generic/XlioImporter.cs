using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Codaxy.Xlio;

namespace Codaxy.Xlio.Generic
{
    public class XlioImporter
    {
        class ImportedColumnInfo<T>
        {
            public int Index { get; set; }
            public ColumnInfo<T> Column { get; set; }
        }

        public static List<T> Import<T>(Sheet sheet, TableInfo<T> t = null, bool skipEmptyRows = true, Cell tableOrigin = null) where T : new()
        {
            if (t == null)
                t = TableInfo<T>.Build();

            if (tableOrigin == null)
                tableOrigin = new Cell();

            var importColumns = new List<ImportedColumnInfo<T>>();
            var headerRow = sheet[tableOrigin.Row];
            foreach (var cell in headerRow)
            {
                if (cell.Key >= tableOrigin.Col && cell.Key <= tableOrigin.Col + t.Columns.Count)
                {
                    var name = cell.Value.Value as String;
                    if (!String.IsNullOrEmpty(name))
                    {
                        var c = t.Columns.FirstOrDefault(a => a.Name == name) ?? t.Columns.FirstOrDefault(a => a.Caption == name);
                        if (c != null && c.Setter != null)
                        {
                            importColumns.Add(new ImportedColumnInfo<T> { Index = cell.Key, Column = c });
                        }
                    }
                }
            }

            List<T> result = new List<T>();

            foreach (var rowData in sheet.Data)
            {
                if (rowData.Key > tableOrigin.Row)
                {
                    bool empty = true;
                    var row = new T();
                    foreach (var ic in importColumns)
                    {
                        var v = rowData.Value[ic.Index].Value;
                        if (v != null)
                        {
                            empty = false;
                            if (ic.Column.ImportConverter != null)
                                v = ic.Column.ImportConverter(v);

                            ic.Column.Setter(row, v);
                        }
                    }
                    if (!empty || !skipEmptyRows)
                        result.Add(row);
                }
            }

            for (var i = tableOrigin.Row + 1; i < sheet.Data.Count; i++)
            {
                if (!skipEmptyRows || sheet.Data[i].Cells.Data.Values.Any(a => a.Value != null))
                {
                    
                }
            }

            return result;            
        }

        public static List<T> Import<T>(String filePath, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            var wb = Xlio.Workbook.Load(filePath);
            return Import(wb, t, skipEmptyRows);
        }

        public static List<T> Import<T>(Stream stream, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            var wb = Xlio.Workbook.LoadFromStream(stream);
            return Import(wb, t, skipEmptyRows);
        }

        public static List<T> Import<T>(Workbook wb, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            var sheet = wb.Sheets[0];
            return Import(sheet, t, skipEmptyRows);
        }
    }
}
