using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class CellDataPair
    {
        public CellDataPair(Cell cell, CellData data)
        {
            Cell = cell;
            Data = data;
        }

        public Cell Cell { get; private set; }
        public CellData Data { get; private set; }
    }
}
