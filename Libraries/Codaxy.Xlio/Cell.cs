using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class Cell 
    {
        public int Row { get; set; }
        public int Col { get; set; }

        public static Cell Parse(String reference)
        {
            return Parse(reference, 0, reference.Length);
        }

        public static Cell Parse(String reference, int start, int count)
        {
            int row = 0;
            int col = 0;
            for (int i = start; i < start + count; i++)
            {
                if (Char.IsDigit(reference[i]))
                    row = 10 * row + (reference[i] - '0');
                else if (Char.IsLetter(reference[i]))
                    col = 26 * col + (Char.ToUpper(reference[i]) - 'A') + 1;
                else
                    throw ExceptionFactory.Argument("Value '{0}' is not a valid cell reference!", reference.Substring(start, count));
            }
            return new Cell { Col = col - 1, Row = row - 1 };
        }

        public static int ParseColumn(String reference)
        {
            return Parse(reference).Col;
        }

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

        public Cell Clone()
        {
            return new Cell { Col = Col, Row = Row };
        }

        public override string ToString()
        {
            return Cell.Format(Row, Col);
        }

        public override int GetHashCode()
        {
            return HashCodeHelper.CalculateHashCode(Row, Col);
        }

        public override bool Equals(object obj)
        {
            Cell o;
            if (!Util.TryCast(obj, out o))
                return false;
            return Row == o.Row && Col == o.Col;
        }

        public static bool operator ==(Cell a, Cell b) {            
            if (Cell.ReferenceEquals(a, b))
                return true;
            if ((object)a==null)
                return false;
            return a.Equals(b);
        }

         public static bool operator !=(Cell a, Cell b) {            
            return !(a==b);
        }
    }
}
