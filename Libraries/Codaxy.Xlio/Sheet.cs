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
            RowBreaks = new List<int>();
            ColumnBreaks = new List<int>();
            Page = new PageSetup();
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

        public SheetRange this[Cell c1, Cell c2]
        {
            get { return new SheetRange(this, new Range { Cell1 = c1, Cell2 = c2 }); }
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
        public double? DefaultColumnWidth { get; set; }
        public bool ShowGridLines { get; set; }
        public Cell ActiveCell { get; set; }

        public SheetColumnCollection Columns { get; private set; }
        public List<int> RowBreaks { get;private set; }
        public List<int> ColumnBreaks { get; private set; }

        public PageSetup Page { get; private set; }

    }

    public enum PageOrientation { Portrait, Landscape }

    public class PageSetup
    {
        public PageSetup()
        {
            Scale = 100;
            Orientation = PageOrientation.Portrait;
        }

        public PageOrientation Orientation { get; set; }
        public PageMargins Margins { get; set; }
        public int Scale { get; set; }
    }

    public class PageMargins
    {
        public double Left { get; set; }
        public double Right { get; set; }
        public double Top { get; set; }
        public double Bottom { get; set; }
        public double Header { get; set; }
        public double Footer { get; set; }
    }

    
}
