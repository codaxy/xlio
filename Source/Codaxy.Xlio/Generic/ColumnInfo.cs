using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Generic
{
    public class ColumnInfo<T>
    {
        public String Caption { get; set; }
        public String Name { get; set; }
        public Func<T, object> Getter { get; set; }
        public Func<object, object> ExportConverter { get; set; }
        public Func<T, CellData> Exporter { get; set; }
        public Func<object, object> ImportConverter { get; set; }
        public Func<CellData, T> Importer { get; set; }
        public Action<T, object> Setter { get; set; }
    }
}
