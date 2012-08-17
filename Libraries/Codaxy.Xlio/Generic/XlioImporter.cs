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

        public static List<T> Import<T>(Sheet sheet, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            if (t == null)
                t = TableInfo<T>.Build();

            var importColumns = new List<ImportedColumnInfo<T>>();
            var firstRow = sheet[0].Data;
            foreach (var cell in firstRow)
            {
                var name = cell.Value.Value as String;
                if (!String.IsNullOrEmpty(name))
                {
                    var c = t.Columns.FirstOrDefault(a => a.Name == name) ?? t.Columns.FirstOrDefault(a => a.Caption == name);
                    if (c != null && c.Setter!=null)
                    {
                        importColumns.Add(new ImportedColumnInfo<T> { Index = cell.Key, Column = c });
                    }
                }
            }

            List<T> result = new List<T>();

            for (var i = 1; i < sheet.Cells.Data.Data.Count; i++)
            {
                if (!skipEmptyRows || sheet.Cells[i].Data.Values.Any(a => a != null))
                {
                    var row = new T();
                    foreach (var ic in importColumns)
                    {
                        var v = sheet.Cells[i, ic.Index].Value;
                        if (ic.Column.ImportConverter != null)
                            v = ic.Column.ImportConverter(v);

                        ic.Column.Setter(row, v);
                    }
                    result.Add(row);
                }
            }

            return result;            
        }

        public static List<T> Import<T>(String filePath, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            var wb = Xlio.Workbook.ReadFile(filePath);
            return Import(wb, t, skipEmptyRows);
        }

        public static List<T> Import<T>(Stream stream, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            var wb = Xlio.Workbook.ReadStream(stream);
            return Import(wb, t, skipEmptyRows);
        }

        public static List<T> Import<T>(Workbook wb, TableInfo<T> t = null, bool skipEmptyRows = true) where T : new()
        {
            var sheet = wb.Sheets[0];
            return Import(sheet, t, skipEmptyRows);
        }
    }
}
