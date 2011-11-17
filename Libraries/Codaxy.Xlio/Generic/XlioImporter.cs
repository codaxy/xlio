using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Codaxy.Xlio;

namespace Codaxy.Xlio.Generic
{
    public class XlioImporter<T> where T:new()
    {
        public static List<T> Import(String filePath, Table<T> t = null)
        {
            var wb = Xlio.Workbook.ReadFile(filePath);
            return Import(wb, t);
        }

        public static List<T> Import(Stream stream, Table<T> t = null)
        {
            var wb = Xlio.Workbook.ReadStream(stream);            
            return Import(wb, t);
        }

        public static List<T> Import(Workbook wb, Table<T> t = null)
        {
            var sheet = wb.Sheets[0];
            return Import(sheet, t);
        }

        class ColumnInfo
        {
            public int Index { get; set; }
            public Table<T>.Column Column { get; set; }
        }

        public static List<T> Import(Sheet sheet, Table<T> t = null)
        {
            if (t == null)
                t = Table<T>.Build();

            var importColumns = new List<ColumnInfo>();
            var firstRow = sheet[0].Data;
            foreach (var cell in firstRow)
            {
                var name = cell.Value.Value as String;
                if (!String.IsNullOrEmpty(name))
                {
                    var c = t.Columns.FirstOrDefault(a => a.Name == name) ?? t.Columns.FirstOrDefault(a => a.Caption == name);
                    if (c != null && c.Setter!=null)
                    {
                        importColumns.Add(new ColumnInfo { Index = cell.Key, Column = c });
                    }
                }
            }

            List<T> result = new List<T>();

            for (var i = 1; i < sheet.Cells.Data.Data.Count; i++)
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

            return result;            
        }
    }
}
