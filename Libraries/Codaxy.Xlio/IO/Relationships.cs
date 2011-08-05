using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio.Model.Opc;

namespace Codaxy.Xlio.IO
{
    class Relationships : IEnumerable<CT_Relationship>
    {
        public const String Sheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        public const String Workbook = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        public const String SharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        public const String Styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        public const String Theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        List<CT_Relationship> data;

        public Relationships() { data = new List<CT_Relationship>(); }

        public String Add(CT_Relationship r)
        {
            r.Id = "rId" + (data.Count + 1);
            data.Add(r);
            return r.Id;
        }

        public IEnumerator<CT_Relationship> GetEnumerator()
        {
            return data.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return data.GetEnumerator();
        }
    }
}
