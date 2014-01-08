using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    //SRP: Manage access to workbook sheet collection
    public class SheetCollection : IEnumerable<Sheet>
    {
        List<Sheet> sheets;
        Dictionary<String, int> sheetNameIndex;        

        public SheetCollection()
        {
            sheets = new List<Sheet>();
            sheetNameIndex = new Dictionary<string, int>();
        }

        public Sheet this[int index] { get { return sheets[index]; } }

        public Sheet this[String sheetName]
        {
            get
            {
                var index = GetSheetIndex(sheetName);
                return sheets[index];
            }
        }

        int GetSheetIndex(String sheetName)
        {
            int index;
            if (sheetNameIndex.TryGetValue(sheetName, out index))
                return index;
            throw ExceptionFactory.Argument("Sheet '{0}' not found.", sheetName);
        }

        public Sheet AddSheet(Sheet newSheet)
        {
            InsertSheet(sheets.Count, newSheet);
            return newSheet;
        }

        public Sheet AddSheet(String sheetName)
        {
            return AddSheet(new Sheet(sheetName));
        }

        public void InsertSheet(int index, Sheet newSheet)
        {
            if (sheetNameIndex.ContainsKey(newSheet.SheetName))
                throw ExceptionFactory.InvalidOperation("Sheet '{0}' already found in workbook.", newSheet.SheetName);            
            sheets.Insert(index, newSheet);
            if (index + 1 == sheets.Count)
                sheetNameIndex.Add(newSheet.SheetName, index);
            else
                RecreateSheetNameIndex();            
        }

        public Sheet RemoveSheet(int index)
        {
            var sheet = sheets[index];
            sheets.RemoveAt(index);
            if (index != sheets.Count)
                RecreateSheetNameIndex();
            else
                sheetNameIndex.Remove(sheet.SheetName);
            return sheet;
        }

        public void RemoveSheet(String sheetName)
        {
            var index = GetSheetIndex(sheetName);
            RemoveSheet(index);
        }

        private void RecreateSheetNameIndex()
        {
            sheetNameIndex.Clear();
            for (var i = 0; i < sheets.Count; i++)
                sheetNameIndex.Add(sheets[i].SheetName, i);
        }

        public IEnumerator<Sheet> GetEnumerator()
        {
            return sheets.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return sheets.GetEnumerator();
        }
    }
}
