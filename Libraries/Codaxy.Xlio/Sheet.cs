using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public partial class Sheet
    {
        public Sheet(String sheetName)
        {
            SheetName = sheetName;
            cells = new Matrix<CellData>();
        }
        
        public string SheetName { get; internal set; }

        Matrix<CellData> cells;
        public Matrix<CellData> Cells { get { return cells; } }

        public CellData this[String location]
        {
            get {
                var cell = Cell.Parse(location);
                return this[cell];
            }
        }        

        public CellData this[Cell cell]
        {
            get { return cells[cell.Row, cell.Col]; }
            set { cells[cell.Row, cell.Col] = value; }
        }

        public SheetRange this[Range range]
        {
            get { return new SheetRange(this, range); }
        }

        public SheetRange this[String c1, String c2]
        {
            get { return new SheetRange(this, new Range { Cell1 = Cell.Parse(c1), Cell2 = Cell.Parse(c2) }); }
        }

        public SheetRange this[int r1, int c1, int r2, int c2]
        {
            get
            {
                return new SheetRange(this, new Range
                {
                    Cell1 = new Cell { Col = c1, Row = r1 },
                    Cell2 = new Cell { Col = c2, Row = r2 }
                });
            }
        }

        public CellData this[int row, int col]
        {
            get { return cells[row, col]; }
            set { cells[row, col] = value; }
        }

        public Array<CellData> this[int row]
        {
            get { return cells[row]; }
            set { cells[row] = value; }
        }
    }
}
