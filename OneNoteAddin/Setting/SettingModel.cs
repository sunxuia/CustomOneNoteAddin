using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace OneNoteAddin.Setting
{
    /// <summary>
    /// setting.xml's model class 
    /// </summary>
    [XmlRoot("setting")]
    public class SettingModel
    {
        /// <summary>
        /// read xml data from file
        /// </summary>
        /// <param name="filePath"></param>
        public SettingModel(string filePath)
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                XmlSerializer xs = new XmlSerializer(typeof(SettingModel));
                var setting = (SettingModel)xs.Deserialize(reader);
                FilePath = filePath;
                CodeStyles = setting.CodeStyles ?? new List<CodeStyleModel>();
                DefaultValues = setting.DefaultValues ?? new List<DefaultValueModel>();
                Tables = setting.Tables ?? new List<TableModel>();
                VSCode = setting.VSCode;
            }
        }

        public SettingModel() { }

        /// <summary>
        /// code style items of combobox in code group
        /// </summary>
        [XmlArray("codeStyles", IsNullable = true)]
        [XmlArrayItem("codeStyle")]
        public List<CodeStyleModel> CodeStyles;

        /// <summary>
        /// setted value of ribbon
        /// </summary>
        [XmlArray("defaultValues", IsNullable = true)]
        [XmlArrayItem("defaultValue")]
        public List<DefaultValueModel> DefaultValues;

        /// <summary>
        /// styles of "add new table"
        /// </summary>
        [XmlArray("tables", IsNullable = true)]
        [XmlArrayItem("table")]
        public List<TableModel> Tables;

        [XmlElement("vsCode")]
        public string VSCode { get; set; }

        /// <summary>
        /// xml file path
        /// </summary>
        [XmlIgnore]
        public string FilePath { get; set; }
        
        /// <summary>
        /// write setting to file
        /// </summary>
        public void WriteXml()
        {
            if (!string.IsNullOrEmpty(FilePath))
            {
                using (StreamWriter writer = new StreamWriter(FilePath))
                {
                    XmlSerializer xs = new XmlSerializer(typeof(SettingModel));
                    xs.Serialize(writer, this);
                    writer.Close();
                }
            }
        }
    }
}
