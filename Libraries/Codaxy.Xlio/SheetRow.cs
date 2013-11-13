using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class SheetRow : IEnumerable<KeyValuePair<int, CellData>>
    {
        readonly Array<CellData> cells = new Array<CellData>();

        public Array<CellData> Cells { get { return cells; } }

        internal CellStyle style;
        public CellStyle Style
        {
            get { return style ?? (style = new CellStyle()); }
            set { style = value; }
        }

        public double? Height { get; set; }
        public bool Collapsed { get; set; }
        public bool Hidden { get; set; }
        public bool Phonetic { get; set; }

        public int LastCol { get { return cells.Data.Count == 0 ? 0 : cells.Data.Last().Key; } }
        public int FirstCol { get { return cells.Data.Count == 0 ? 0 : cells.Data.First().Key; } }
        public int Count { get { return cells.Data.Count; } }        

        public CellData this[int index]
        {
            get { return cells[index]; }
            set { cells[index] = value; }
        }

        IEnumerator<KeyValuePair<int, CellData>> IEnumerable<KeyValuePair<int, CellData>>.GetEnumerator()
        {
            return cells.Data.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return cells.Data.GetEnumerator();
        }
    }

    public class SheetData : IEnumerable<KeyValuePair<int, SheetRow>>
    {
        readonly Array<SheetRow> rows = new Array<SheetRow>();

        public Array<SheetRow> Rows { get { return rows; } }

        public CellData this[int row, int col]
        {
            get
            {
                return rows[row][col];
            }
            set
            {
                rows[row][col] = value;
            }
        }

        public SheetRow this[int row]
        {
            get { return rows[row]; }
            set { rows[row] = value; }
        }

        public int Count { get { return rows.Count; } }
        public int LastRow { get { return rows.Count == 0 ? 0 : rows.Data.Last().Key; } }
        public int FirstRow { get { return rows.Count == 0 ? 0 : rows.Data.First().Key; } }

        public int LastCol { get { return rows.Data.Values.Max(a => a.LastCol); } }
        public int FirstCol { get { return rows.Data.Values.Min(a => a.FirstCol); } }       

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return rows.Data.GetEnumerator();
        }

        IEnumerator<KeyValuePair<int, SheetRow>> IEnumerable<KeyValuePair<int, SheetRow>>.GetEnumerator()
        {
            return rows.Data.GetEnumerator();
        }
    }
}
