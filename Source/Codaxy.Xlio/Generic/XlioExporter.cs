using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Codaxy.Xlio;

namespace Codaxy.Xlio.Generic
{
    public class XlioExporter
    {
        public static void Export<T>(IEnumerable<T> data, String filePath, TableInfo<T> t = null) where T:new()
        {
            using (var fs = File.Create(filePath))
                Export(data, fs, t);
        }

        public static void Export<T>(IEnumerable<T> data, Stream stream, TableInfo<T> t = null) where T:new()
        {
            var wb = new Workbook();
            var sheet = new Sheet("Sheet 1");
            Export(data, sheet, t);
            wb.Sheets.AddSheet(sheet);
            wb.SaveToStream(stream, IO.XlsxFileWriterOptions.AutoFit);
        }

        public static void Export<T>(IEnumerable<T> data, Sheet sheet, TableInfo<T> t = null, Cell tableOrigin = null) where T : new()
        {
            if (t == null)
                t = TableInfo<T>.Build();

            if (tableOrigin == null)
                tableOrigin = new Cell(0, 0);

            var exportColumns = t.Columns.Where(a => a.Getter != null).ToArray();



            for (var i = 0; i < exportColumns.Length; i++)
            {
                sheet[tableOrigin.Row, tableOrigin.Col + i].Value = exportColumns[i].Caption ?? exportColumns[i].Name;
                sheet[tableOrigin.Row, tableOrigin.Col + i].Style = exportColumns[i].HeaderStyle;

                if (exportColumns[i].ExportHeaderFormatter != null)
                    exportColumns[i].ExportHeaderFormatter(exportColumns[i], sheet[tableOrigin.Row, tableOrigin.Col + i]);

                if (exportColumns[i].ExportColumnFormatter != null)
                    exportColumns[i].ExportColumnFormatter(sheet.Columns[tableOrigin.Col + i]);
            }

            int row = 1;

            foreach (var record in data)
            {
                for (var i = 0; i < exportColumns.Length; i++)
                {
                    if (exportColumns[i].Exporter != null)
                    {
                        var cell = exportColumns[i].Exporter(record);
                        sheet[tableOrigin.Row + row, tableOrigin.Col + i] = cell;
                    }
                    else
                    {
                        var value = exportColumns[i].Getter(record);
                        if (exportColumns[i].ExportConverter != null)
                            value = exportColumns[i].ExportConverter(value);
                        sheet[tableOrigin.Row + row, tableOrigin.Col + i].Value = value;
                        if (exportColumns[i].ExportFormatter != null)
                            exportColumns[i].ExportFormatter(sheet[row, i], record, i, value);
                    }
                }
                ++row;
            }
        }
    }
}
