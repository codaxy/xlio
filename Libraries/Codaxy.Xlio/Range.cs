using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class Range 
    {
        public Cell Cell1 { get; set;}
        public Cell Cell2 { get; set; }

        public Range() { Cell1 = new Cell(); Cell2 = new Cell(); }

        public int Width { get { return Cell2.Col - Cell1.Col + 1; } }
        public int Height { get { return Cell2.Row - Cell1.Row + 1; } }

        public bool Contains(int row, int col)
        {
            return Cell1.Col <= col && col <= Cell2.Col && Cell1.Row <= row && row <= Cell2.Row;
        }

        public Cell GetCell(int index)
        {
            return new Cell
            {
                Row = Cell1.Row + index / Width,
                Col = Cell1.Col + index % Width
            };
        }

        public Range GetRow(int index)
        {
            return new Range { Cell1 = new Cell { Col = Cell1.Col, Row = Cell1.Row + index }, Cell2 = new Cell { Col = Cell2.Col, Row = Cell1.Row + index } };
        }

        public Range GetRows(int firstRow, int lastRow)
        {
            return new Range { Cell1 = new Cell { Col = Cell1.Col, Row = Cell1.Row + firstRow }, Cell2 = new Cell { Col = Cell2.Col, Row = Cell1.Row + lastRow } };
        }

        public Range GetColumn(int index)
        {
            return new Range { Cell1 = new Cell { Col = Cell1.Col + index, Row = Cell1.Row }, Cell2 = new Cell { Col = Cell1.Col + index, Row = Cell2.Row } };
        }

        public Range GetColumns(int firstCol, int lastCol)
        {
            return new Range { Cell1 = new Cell { Col = Cell1.Col + firstCol, Row = Cell1.Row }, Cell2 = new Cell { Col = Cell1.Col + lastCol, Row = Cell2.Row } };
        }

        public bool Contains(Cell cell) { return Contains(cell.Row, cell.Col); }

        public static Range Parse(string range)
        {
            int colon = range.IndexOf(':');
            if (colon != -1)
            {
                var c1 = Cell.Parse(range, 0, colon);
                var c2 = Cell.Parse(range, colon + 1, range.Length - colon - 1);
                return new Range
                {
                    Cell1 = c1, Cell2 = c2
                };
            }
            else
            {
                var c = Cell.Parse(range);
                return new Range
                {
                    Cell1 = c,
                    Cell2 = c.Clone()
                };
            }
        }

        public override string ToString()
        {
            return Cell1.ToString() + ":" + Cell2.ToString();
        }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Cell1, Cell2);
        }

        public override bool Equals(object obj)
        {
            Range o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Cell1 == o.Cell1 && Cell2 == o.Cell2;
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static bool operator ==(Range a, Range b)
        {
            if (Range.ReferenceEquals(a, b))
                return true;
            if ((object)a == null)
                return false;
            return a.Equals(b);
        }

        public static bool operator !=(Range a, Range b)
        {
            return !(a == b);
        }
    }
}
