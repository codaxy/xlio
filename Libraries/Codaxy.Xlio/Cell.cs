using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    /// <summary>
    /// Cell location e.g. A1
    /// </summary>
    public class Cell 
    {
        /// <summary>
        /// Zero indexed row number
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Zero indexed column number
        /// </summary>
        public int Col { get; set; }


        /// <summary>
        /// Convert to cell from its string representaion
        /// </summary>
        /// <param name="reference">Textual cell reference, e.g. A1</param>
        /// <returns></returns>
        public static Cell Parse(String reference)
        {
            return Parse(reference, 0, reference.Length);
        }

        public Cell Offset(int rows, int cols)
        {
            return new Cell { Row = Row + rows, Col = Col + cols };
        }

        /// <summary>
        /// Parse cell from the part of the string
        /// </summary>
        /// <param name="reference">Textual cell reference, e.g. A1</param>
        /// <param name="offset"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static Cell Parse(String reference, int offset, int count)
        {
            int row = 0;
            int col = 0;
            for (int i = offset; i < offset + count; i++)
            {
                if (Char.IsDigit(reference[i]))
                    row = 10 * row + (reference[i] - '0');
                else if (Char.IsLetter(reference[i]))
                    col = 26 * col + (Char.ToUpper(reference[i]) - 'A') + 1;
                else
                    throw ExceptionFactory.Argument("Value '{0}' is not a valid cell reference!", reference.Substring(offset, count));
            }
            return new Cell { Col = col - 1, Row = row - 1 };
        }

        /// <summary>
        /// Convert string column name to column offset number
        /// </summary>
        /// <param name="reference"></param>
        /// <returns></returns>
        public static int ParseColumn(String reference)
        {
            return Parse(reference).Col;
        }

        /// <summary>
        /// Format cell e.g. (0, 0) => A1
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static string Format(int row, int col)
        {
            Stack<char> stack = new Stack<char>();
            do
            {
                stack.Push((Char)('A' + (col % 26)));
                col = col / 26 - 1;
            } while (col >= 0);
            return new String(stack.ToArray()) + (row + 1).ToString();
        }

        /// <summary>
        /// Clones the cell.
        /// </summary>
        /// <returns></returns>
        public Cell Clone()
        {
            return new Cell { Col = Col, Row = Row };
        }

        /// <summary>
        /// Returns cell's string representation, e.g. A1.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Cell.Format(Row, Col);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Row, Col);
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object" /> is equal to this instance.
        /// </summary>
        /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
        /// <returns>
        ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
        /// </returns>
        public override bool Equals(object obj)
        {
            Cell o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Row == o.Row && Col == o.Col;
        }

        /// <summary>
        /// Operator ==
        /// </summary>
        /// <param name="a">A.</param>
        /// <param name="b">The b.</param>
        /// <returns></returns>
        public static bool operator ==(Cell a, Cell b)
        {
            if (Cell.ReferenceEquals(a, b))
                return true;
            if ((object)a == null)
                return false;
            return a.Equals(b);
        }

        /// <summary>
        /// Operator !=
        /// </summary>
        /// <param name="a">A.</param>
        /// <param name="b">The b.</param>
        /// <returns></returns>
        public static bool operator !=(Cell a, Cell b)
        {
            return !(a == b);
        }
    }
}
