using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace OneNoteAddin.Setting
{
    /// <summary>
    /// insert table's style model in setting.xml
    /// </summary>
    public class TableModel
    {
        [XmlAttribute("label")]
        public string Label { get; set; }

        [XmlAttribute("size")]
        public string Size { get; set; }

        [XmlAttribute("row")]
        public int Row { get; set; }

        [XmlAttribute("column")]
        public int Column { get; set; }

        [XmlAttribute("foreColor")]
        public string ForeColor { get; set; }

        [XmlAttribute("backColor")]
        public string BackColor { get; set; }

        [XmlAttribute("headInLeft")]
        public bool HeadInLeft { get; set; }
    }
}
