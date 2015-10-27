using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Codaxy.Xlio
{
    namespace Model.Oxml
    {
        public partial class CT_Rst
        {
            [XmlAttribute("xml:space")]
            public String Space
            {
                get; set;
            }
        }
    }
}
