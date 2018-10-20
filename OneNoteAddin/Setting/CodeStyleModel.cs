using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace OneNoteAddin.Setting
{
    public class CodeStyleModel
    {
        [XmlAttribute("label")]
        public string Label { get; set; }

        [XmlAttribute("comment")]
        public string Comment { get; set; }
    }
}
