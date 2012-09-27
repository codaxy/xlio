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
            data = new SheetData();
            Columns = new SheetColumnCollection();
        }
        
        public string SheetName { get; internal set; }

        SheetData data;
        public SheetData Data { get { return data; } }

        public CellData this[String location]
        {
            get {
                var cell = Cell.Parse(location);
                return this[cell];
            }
        }        

        public CellData this[Cell cell]
        {
            get { return data[cell.Row, cell.Col]; }
            set { data[cell.Row, cell.Col] = value; }
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
            get { return data[row, col]; }
            set { data[row, col] = value; }
        }

        public SheetRow this[int row]
        {
            get { return data[row]; }
            set { data[row] = value; }
        }

        public double? DefaultRowHeight { get; set; }
        public bool ShowGridLines { get; set; }
        public Cell ActiveCell { get; set; }

        public SheetColumnCollection Columns { get; private set; }
    }
}
