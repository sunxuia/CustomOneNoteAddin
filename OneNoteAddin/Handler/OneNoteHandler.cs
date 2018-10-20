using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNoteAddin.Handler
{
    public class OneNotePageHandler
    {
        private Application app;

        private XDocument doc;

        public OneNotePageHandler(Application app)
        {
            this.app = app;
            string pageId = app.Windows.CurrentWindow.CurrentPageId;
            string xml;
            app.GetPageContent(pageId, out xml, PageInfo.piAll);
            doc = XDocument.Parse(xml);
        }

        public string GetSelectedText()
        {
            return GetInnerText(Descendants(doc.Root, "T")
                .Where(n => HasAttributeValue(n, "selected", "all")));
        }

        private IEnumerable<XElement> Descendants(XElement node, string childNodeName)
        {
            return node.Descendants(doc.Root.Name.Namespace + childNodeName);
        }

        private bool HasAttributeValue(XElement element, string attrName, params string[] candidates)
        {
            var attr = element.Attribute(attrName);
            if (attr != null)
            {
                foreach (var candidate in candidates)
                {
                    if (attr.Value == candidate)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public string GetInnerText(IEnumerable<XElement> elements)
        {
            StringBuilder sb = new StringBuilder();
            elements.Select(n =>
            {
                var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(n.Value);
                return HttpUtility.HtmlDecode(htmlDocument.DocumentNode.InnerText);
            })
            .ToList()
            .ForEach(s => sb.Append(s).Append("\n"));
            return sb.ToString();
        }

        public IEnumerable<XElement> GetSelectedElements(string elementName)
        {
            return Descendants(doc.Root, elementName)
                .Where(n => HasAttributeValue(n, "selected", "partial", "all"));
        }

        public void EnumTextHtml(IEnumerable<XElement> elements, Action<HtmlAgilityPack.HtmlDocument> action)
        {
            foreach (var element in elements)
            {
                foreach (var t in Descendants(element, "T"))
                {
                    var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                    htmlDocument.LoadHtml(t.Value);
                    action.Invoke(htmlDocument);
                    t.Value = htmlDocument.Text;
                }
            }
        }

        public void EnumPageTextHtml(Action<HtmlAgilityPack.HtmlDocument> action)
        {
            EnumTextHtml(Enumerable.Repeat(doc.Root, 1), action);
        }

        public XElement GetCursorElement(string elementName)
        {
            return Descendants(doc.Root, elementName)
                .Where(n => HasAttributeValue(n, "selected", "partial", "all"))
                .LastOrDefault();
        }

        public string GetInnerText(XElement element)
        {
            return GetInnerText(Descendants(element, "T"));
        }

        public void SetAttributeValue(IEnumerable<XElement> elements,
            string attrName, AttributeValueModifier modifier)
        {
            foreach (var element in elements)
            {
                var attr = element.Attribute(attrName);
                bool exist = attr != null;
                string value = attr?.Value;
                modifier.Invoke(ref exist, ref value);
                if (attr != null && !exist)
                {
                    attr.Remove();
                }
                else if (attr != null && attr.Value != value)
                {
                    element.SetAttributeValue(attrName, value);
                }
            }
        }

        public delegate void AttributeValueModifier(ref bool attrExist, ref string value);

        public void SetPageAttributes(string attrName, AttributeValueModifier modifier)
        {
            SetAttributeValue(doc.Root.Descendants(), attrName, modifier);
        }

        public void SetQuickStyleDef(string attrName, string attrValue)
        {
            var def = Descendants(doc.Root, "QuickStyleDef").FirstOrDefault();
            if (def != null)
            {
                def.SetAttributeValue(attrName, attrValue);
            }
        }

        public void Save()
        {
            app.UpdatePageContent(doc.ToString(), DateTime.MinValue);
        }
    }
}
