using Extensibility;
using Microsoft.Office.Core;
using OneNoteAddin.Setting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace OneNoteAddin
{
    /// <summary>
    /// custom process while load this addin
    /// </summary>
    public partial class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        private SettingModel setting;

        #region ribbon.xml
        public string GetCustomUI(string ribbonID)
        {
            try
            {
                //System.Diagnostics.Debugger.Launch();
                LoadSetting();

                string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)), "ribbon.xml");
                XmlDocument doc = new XmlDocument();
                doc.Load(filePath);
                EditXml(doc);

                return GetXml(doc);
            }
            catch (Exception err)
            {
                MessageBox.Show("Error : " + err.ToString());
                throw;
            }
        }

        private void LoadSetting()
        {
            string filePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)), "setting.xml");
            setting = new SettingModel(filePath);
        }

        private void EditXml(XmlDocument doc)
        {
            XmlElement root = doc.DocumentElement;
            string xmlns = root.NamespaceURI;
            XmlNamespaceManager namespaceManager = new XmlNamespaceManager(doc.NameTable);
            namespaceManager.AddNamespace("r", xmlns);

            // code style combobox
            var codeStyleNode = root.SelectSingleNode("//r:comboBox[@id='cmbStyle']", namespaceManager);
            if (codeStyleNode != null)
            {
                AddCodeStyles(codeStyleNode as XmlElement, doc);
            }

            // font combobox
            foreach (XmlNode node in root.SelectNodes("//r:comboBox[@imageMso='FontColorGallery']", namespaceManager))
            {
                AddFontItems(node as XmlElement, doc);
            }

            // insert table bottons
            XmlElement grpTable = root.SelectSingleNode("//r:group[@id='grpTable']", namespaceManager) as XmlElement;
            AddTables(grpTable, doc);
        }

        private void AddFontItems(XmlElement node, XmlDocument doc)
        {
            string nodeId = node.Attributes["id"].Value;
            var fonts = new System.Drawing.Text.InstalledFontCollection();
            for (int i = 0; i < fonts.Families.Length; i++)
            {
                if (!string.IsNullOrEmpty(fonts.Families[i].Name))
                {
                    XmlElement item = doc.CreateElement("item", node.NamespaceURI);
                    item.SetAttribute("id", "__" + nodeId + "_" + i);
                    item.SetAttribute("label", fonts.Families[i].Name);
                    node.AppendChild(item);
                }
            }
        }

        private void AddCodeStyles(XmlElement node, XmlDocument doc)
        {
            string nodeId = node.Attributes["id"].Value;
            int i = 0;
            foreach (var style in setting.CodeStyles)
            {
                XmlElement item = doc.CreateElement("item", node.NamespaceURI);
                item.SetAttribute("id", "__" + nodeId + "_" + i++);
                item.SetAttribute("label", style.Label); ;
                node.AppendChild(item);
            }
        }

        private void AddTables(XmlElement node, XmlDocument doc)
        {
            string nodeId = node.Attributes["id"].Value;
            for (int i = 0; i < setting.Tables.Count; i++)
            {
                var table = setting.Tables[i];
                XmlElement btn = doc.CreateElement("button", node.NamespaceURI);
                btn.SetAttribute("id", "__" + nodeId + "_" + i);
                btn.SetAttribute("imageMso", "AdpDiagramAddTable");
                btn.SetAttribute("label", table.Label);
                if (table.Size != null)
                {
                    btn.SetAttribute("size", table.Size);
                }
                btn.SetAttribute("onAction", "OnInsertTableClick");
                node.AppendChild(btn);
            }
        }

        private string GetXml(XmlDocument doc)
        {
            MemoryStream stream = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(stream, null);
            writer.Formatting = Formatting.Indented;
            doc.Save(writer);
            stream.Position = 0;
            StreamReader sr = new StreamReader(stream, Encoding.UTF8);
            string xml = sr.ReadToEnd();
            sr.Close();
            stream.Close();
            return xml;
        }
        #endregion

        public string GetText(IRibbonControl control)
        {
            string id = GetControlId(control);
            return GetDefaultValue(id);
        }

        private string GetControlId(IRibbonControl control)
        {
            return ((dynamic)control).Id;
        }

        private string GetDefaultValue(string id)
        {
            foreach (var defaultValue in setting.DefaultValues)
            {
                if (defaultValue.Id == id)
                {
                    return defaultValue.Value;
                }
            }
            return "";
        }

        public IStream GetImage(string imageName)
        {
            MemoryStream imageStream = new MemoryStream();
            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            var prop = typeof(Properties.Resources).GetProperty(imageName ?? "", flags);
            if (prop != null)
            {
                (prop.GetValue(null, null) as Image).Save(imageStream, ImageFormat.Png);
            }
            else
            {
                Properties.Resources.DefaultImage.Save(imageStream, ImageFormat.Png);
            }
            return new CCOMStreamWrapper(imageStream);
        }

        private void SaveSetting()
        {
            setting?.WriteXml();
        }
    }
}
