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
        public Range(Cell cell1, Cell cell2)
        {
            Cell1 = cell1;
            Cell2 = cell2;
        }

        public Range(String cell1, String cell2)
        {
            Cell1 = Cell.Parse(cell1);
            Cell2 = Cell.Parse(cell2);
        }

        public Range(int r1, int c1, int r2, int c2)
        {
            Cell1 = new Cell(r1, c1);
            Cell2 = new Cell(r2, c2);
        }

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

        public int Area { get { return Width * Height; } }

        public static string Format(Range range, bool absolute = false)
        {
            if (absolute)
            {
                var cell1 = range.Cell1;
                var cell2 = range.Cell2;
                var cell1Str = Cell.Format(cell1.Row, cell1.Col, true, true);
                var cell2Str = Cell.Format(cell2.Row, cell2.Col, true, true);
                return cell1Str + ":" + cell2Str;
            }
            return range.ToString();
        }

        public static string Format(Cell cell1, Cell cell2)
        {
            return cell1.ToString() + ":" + cell1.ToString();
        }

        public static string Format(String cell1, String cell2)
        {
            return cell1 + ":" + cell2;
        }

        public static string Format(int r1, int c1, int r2, int c2)
        {
            return Format(new Cell(r1, c1), new Cell(r2, c2));
        }
    }
}
