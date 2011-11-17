using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Codaxy.Xlio;

namespace Codaxy.Xlio.Generic
{
    public class XlioExporter<T> where T:new()
    {
        public static void Export(IEnumerable<T> data, String filePath, Table<T> t = null)
        {
            using (var fs = File.Create(filePath))
                Export(data, fs, t);
        }

        public static void Export(IEnumerable<T> data, Stream stream, Table<T> t = null)
        {
            var wb = new Workbook();
            var sheet = new Sheet("Sheet 1");
            Export(data, sheet, t);
            wb.Sheets.AddSheet(sheet);
            wb.SaveToStream(stream);
        }

        public static void Export(IEnumerable<T> data, Sheet sheet, Table<T> t = null)
        {
            if (t == null)
                t = Table<T>.Build();

            var exportColumns = t.Columns.Where(a => a.Getter != null).ToArray();

            for (var i = 0; i < exportColumns.Length; i++)
            {
                sheet[0, i].Value = exportColumns[i].Caption ?? exportColumns[i].Name;
                sheet[0, i].Style.Fill = new CellFill { Pattern = FillPattern.Solid, Foreground = new Color(255, 0xAD, 0xD8, 0xE6) };
            }

            int row = 1;

            foreach (var record in data)
            {
                for (var i = 0; i < exportColumns.Length; i++)
                {
                    if (exportColumns[i].Exporter != null)
                    {
                        var cell = exportColumns[i].Exporter(record);
                        sheet[row, i] = cell;
                    }
                    else
                    {
                        var value = exportColumns[i].Getter(record);
                        if (exportColumns[i].ExportConverter != null)
                            value = exportColumns[i].ExportConverter(value);
                        sheet[row, i].Value = value;
                    }
                }
                ++row;
            }
        }
    }
}
