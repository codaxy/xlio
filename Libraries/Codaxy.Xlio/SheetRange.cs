using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class SheetRange
    {
        public Range Range { get; private set; }
        public Sheet Sheet { get; private set; }

        public SheetRange(Sheet sheet, Range range)
        {
            Sheet = sheet;
            Range = range;
        }

        public CellData Merge()
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                    rowRange[col].MergedRange = Range;
            }

            return this[0];
        }

        public SheetRange UnMerge()
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                    if (rowRange[col].IsMerged)
                    {
                        var merge = rowRange[col].MergedRange;
                        for (var row1 = merge.Cell1.Row; row1 <= merge.Cell2.Row; row1++)
                            for (var col1 = merge.Cell1.Col; col1 <= merge.Cell2.Col; col1++)
                                Sheet[col1, row1].MergedRange = null;
                    }
            }

            return this;
        }

        public SheetRange SetOutsideBorder(BorderEdge style)
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                Sheet[row, Range.Cell1.Col].Style.Border.Left = style;
                Sheet[row, Range.Cell2.Col].Style.Border.Right = style;
            }

            for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
            {
                Sheet[Range.Cell1.Row, col].Style.Border.Top = style;
                Sheet[Range.Cell2.Row, col].Style.Border.Bottom = style;
            }

            return this;
        }

        public SheetRange SetInsideBorder(BorderEdge style)
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                {
                    if (col > Range.Cell1.Col)
                        Sheet[row, col].Style.Border.Left = style;

                    if (col < Range.Cell2.Col)
                        Sheet[row, col].Style.Border.Right = style;

                    if (row > Range.Cell1.Row)
                        Sheet[row, col].Style.Border.Top = style;

                    if (row < Range.Cell2.Row)
                        Sheet[row, col].Style.Border.Bottom = style;
                }

            return this;
        }

        public SheetRange ClearData()
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                {
                    var cell = rowRange[col];
                    cell.Value = null;
                    cell.Formula = null;
                    cell.SharedFormula = null;
                }
            }

            return this;
        }

        public SheetRange ClearFormatting()
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                {
                    var cell = rowRange[col];
                    cell.Style = null;
                }
            }

            return this;
        }

        public SheetRange SetBorder(BorderEdge style)
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                {
                    var border = rowRange[col].Style.Border;
                    border.Left = style;
                    border.Right = style;
                    border.Top = style;
                    border.Bottom = style;
                }
            }

            return this;
        }

        public IEnumerable<CellData> EnumerateCells()
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                    yield return rowRange[col];
            }
        }

        public SheetRange SetStyle(CellStyle style)
        {
            foreach (var cell in EnumerateCells())
                cell.Style = style;
            return this;
        }

        public SheetRange ApplyStyle(CellStyle style)
        {
            foreach (var cell in EnumerateCells())
                cell.Style.Apply(style);
            return this;
        }

        public CellData this[int relativeRow, int relativeCol]
        {
            get { return Sheet[Range.Cell1.Row + relativeRow, Range.Cell1.Col + relativeCol]; }
            set { Sheet[Range.Cell1.Row + relativeRow, Range.Cell1.Col + relativeCol] = value; }
        }

        public SheetRange this[int relativeRow1, int relativeCol1, int relativeRow2, int relativeCol2]
        {
            get { return Sheet[Range.Cell1.Offset(relativeRow1, relativeCol1), Range.Cell1.Offset(relativeRow2, relativeCol2)]; }
        }

        public CellData this[int index]
        {
            get { return Sheet[Range.GetCell(index)]; }
            set { Sheet[Range.GetCell(index)] = value; }
        }

        public SheetRange GetRow(int index)
        {
            return Sheet[Range.GetRow(index)];
        }

        public SheetRange GetColumn(int index)
        {
            return Sheet[Range.GetColumn(index)];
        }

        


    }
}
