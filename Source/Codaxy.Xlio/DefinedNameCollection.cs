using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class DefinedNameCollection : IEnumerable<DefinedName>
    {
        private Dictionary<string, DefinedName> definedNames;

        public DefinedNameCollection()
        {
            definedNames = new Dictionary<string, DefinedName>();
        }

        public DefinedName AddDefinedName(DefinedName definedName)
        {
            definedNames.Add(definedName.Name, definedName);
            return definedName;
        }

        public DefinedName AddCell(String name, Cell cell, Sheet sheet)
        {

            DefinedName dn = new DefinedName
            {
                Name = name,
                Value = GetSheetName(sheet) + "!" + Cell.Format(cell.Row, cell.Col, true, true)
            };
            definedNames.Add(name, dn);
            return dn;
        }

        public DefinedName AddRange(String name, Range range, Sheet sheet)
        {

            DefinedName dn = new DefinedName
            {
                Name = name,
                Value = GetSheetName(sheet) + "!" + Range.Format(range, true)
            };
            definedNames.Add(name, dn);
            return dn;
        }

        public DefinedName AddValue(String name, String value)
        {
            DefinedName dn = new DefinedName
            {
                Name = name,
                Value = value
            };
            definedNames.Add(name, dn);
            return dn;
        }

        public void Remove(String name)
        {
            definedNames.Remove(name);
        }

        public DefinedName this[String name]
        {
            get
            {
                return definedNames[name];
            }
            set
            {
                definedNames[name] = value;
            }
        }

        int Count { get { return definedNames.Count; } }

        public IEnumerator<DefinedName> GetEnumerator()
        {
            return definedNames.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return definedNames.Values.GetEnumerator();
        }
        private String GetSheetName(Sheet sheet)
        {
            if (sheet.SheetName.Contains(" "))
                return "'" + sheet.SheetName + "'";
            return sheet.SheetName;
        }
    }
}
