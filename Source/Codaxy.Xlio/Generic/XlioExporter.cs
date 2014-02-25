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

        public static void Export<T>(IEnumerable<T> data, Sheet sheet, TableInfo<T> t = null) where T : new()
        {
            if (t == null)
                t = TableInfo<T>.Build();

            var exportColumns = t.Columns.Where(a => a.Getter != null).ToArray();

            var headerStyle = new CellStyle
            {
                Fill = new CellFill { Pattern = FillPattern.Solid, Foreground = new Color(255, 0x46, 0x82, 0xB4) },
                Font = new CellFont { Color = new Color(255, 255, 255, 255), Bold = true }
            };

            for (var i = 0; i < exportColumns.Length; i++)
            {
                sheet[0, i].Value = exportColumns[i].Caption ?? exportColumns[i].Name;
                sheet[0, i].Style = headerStyle;
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
                        if (exportColumns[i].ExportFormatter != null)
                            exportColumns[i].ExportFormatter(sheet[row, i], record, i, value);
                    }
                }
                ++row;
            }
        }
    }
}
