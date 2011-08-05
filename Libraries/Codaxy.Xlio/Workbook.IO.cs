using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Codaxy.Xlio.IO;

namespace Codaxy.Xlio
{
    public partial class Workbook
    {
        public static Workbook ReadFile(String path)
        {
            using (var fs = File.OpenRead(path))
                return ReadStream(fs);
        }

        public static Workbook ReadStream(Stream fs)
        {
            using (var fr = new XlsxFileReader(fs))
            {
                return fr.Read();
            }
        }

        public void Save(String path) { Save(path, XlsxFileWriterOptions.None); }
        public void Save(String path, XlsxFileWriterOptions options)
        {
            using (var fs = File.Create(path))
                SaveToStream(fs, options);
        }

        public void SaveToStream(Stream fs) { SaveToStream(fs, XlsxFileWriterOptions.None); }
        public void SaveToStream(Stream fs, XlsxFileWriterOptions options)
        {
            using (var xw = new XlsxFileWriter(fs) { Options = options })
                xw.Write(this);
        }
    }
}
