using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace OneNoteAddin.Setting
{
    /// <summary>
    /// default value model in setting.xml
    /// </summary>
    public class DefaultValueModel
    {
        [XmlAttribute("id")]
        public string Id { get; set; }

        [XmlAttribute("value")]
        public string Value { get; set; }
    }
}
