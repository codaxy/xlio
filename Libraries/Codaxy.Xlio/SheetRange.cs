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

        public void Merge()
        {
            for (var row = Range.Cell1.Row; row <= Range.Cell2.Row; row++)
            {
                var rowRange = Sheet[row];
                for (var col = Range.Cell1.Col; col <= Range.Cell2.Col; col++)
                    rowRange[col].MergedRange = Range;
            }
        }

        public void UnMerge()
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
        }

        public void SetOutsideBorder(BorderEdge style)
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
        }

        public void ClearData()
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
        }

        public void ClearFormatting()
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
        }

        public void SetBorder(BorderEdge style)
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
        }

        public CellData this[int relativeRow, int relativeCol]
        {
            get { return Sheet[Range.Cell1.Row + relativeRow, Range.Cell1.Col + relativeCol]; }
            set { Sheet[Range.Cell1.Row + relativeRow, Range.Cell1.Col + relativeCol] = value; }
        }

        public CellData this[int index]
        {
            get { return Sheet[GetCell(index)]; }
            set { Sheet[GetCell(index)] = value; }
        }

        public Cell GetCell(int index)
        {
            return new Cell
            {
                Row = Range.Cell1.Row + index / Range.Width,
                Col = Range.Cell1.Col + index % Range.Width
            };
        }
    }
}
