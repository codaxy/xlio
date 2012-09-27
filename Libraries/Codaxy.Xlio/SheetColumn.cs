using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class SheetColumn
    { 
        public double? Width { get; set; }

        internal CellStyle style;
        public CellStyle DefaultStyle
        {
            get { return style ?? (style = new CellStyle()); }
            set { style = value; }
        }

        public bool BestFit { get; set; }
        public bool Phonetic { get; set; }
        public byte OutlineLevel { get; set; }
        public bool Hidden { get; set; }
    }

    class SheetColumnRange
    {
        //Zero indexed column A - 0
        public int FirstColumn { get; set; }

        //Zero indexed column A - 0
        public int LastColumn { get; set; }

        public SheetColumn Data { get; set; }
    }

    public class SheetColumnCollection
    {
        internal readonly LinkedList<SheetColumnRange> data = new LinkedList<SheetColumnRange>();

        public SheetColumn this[int col]
        {
            get
            {
                var node = GetNode(col);
                return node.Value.Data;
            }
            set
            {
                var node = GetNode(col);
                node.Value.Data = value;
            }
        }

        public SheetColumn this[String col]
        {
            get { return this[Cell.ParseColumn(col)]; }
            set { this[Cell.ParseColumn(col)] = value; }
        }

        public SheetColumn this[int firstCol, int lastCol]
        {
            set { SetRange(firstCol, lastCol, value); }
        }

        public SheetColumn this[String firstCol, String lastCol]
        {
            set { SetRange(Cell.ParseColumn(firstCol), Cell.ParseColumn(lastCol), value); }
        }

        public void SetRange(int firstCol, int lastCol, SheetColumn column)
        {
            if (lastCol < firstCol)
                throw new InvalidOperationException("Invalid range.");

            var node = GetNode(firstCol);
            node.Value.Data = column;
            node.Value.LastColumn = lastCol;

            var next = node.Next;
            while (next != null && next.Value.FirstColumn <= lastCol)
            {
                var at = next;
                next = at.Next;

                if (at.Value.LastColumn <= lastCol)
                    data.Remove(at);
                else
                    at.Value.FirstColumn = lastCol + 1;
            }
        }

        LinkedListNode<SheetColumnRange> GetNode(int col)
        {
            var at = data.First;
            while (at != null && at.Value.FirstColumn < col)
                at = at.Next;

            //at the end
            if (at == null || at.Value.LastColumn < col)
            {
                var c = new SheetColumnRange { FirstColumn = col, LastColumn = col, Data = new SheetColumn() };
                return data.AddLast(c);
            }            

            //between nodes
            if (at.Value.FirstColumn > col)
                return data.AddBefore(at, new SheetColumnRange { FirstColumn = col, LastColumn = col, Data = new SheetColumn() });

            //overlap 

            if (col == at.Value.LastColumn)
                return at;

            if (col < at.Value.LastColumn)
                data.AddAfter(at, new SheetColumnRange { FirstColumn = col + 1, LastColumn = at.Value.LastColumn, Data = at.Value.Data });

            var result = data.AddAfter(at, new SheetColumnRange { FirstColumn = col, LastColumn = col, Data = new SheetColumn() });

            if (col > at.Value.FirstColumn)
                at.Value.LastColumn = col - 1;
            else
                data.Remove(at);

            return result;
        }

        /// <summary>
        /// Used in loading data
        /// </summary>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="sheetColumn"></param>
        internal void AppendRange(int firstCol, int lastCol, SheetColumn sheetColumn)
        {
            data.AddLast(new SheetColumnRange { FirstColumn = firstCol, LastColumn = lastCol, Data = sheetColumn });
        }
    }
}
