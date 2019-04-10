using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using Forms = System.Windows.Forms;
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
            return GetInnerHtmlText(GetSelectedTextElement());
        }

        public IEnumerable<XElement> GetSelectedTextElement()
        {
            return Descendants(doc.Root, "T")
                .Where(n => HasAttributeValue(n, "selected", "all"));
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

        private string GetInnerHtmlText(IEnumerable<XElement> elements)
        {
            StringBuilder sb = new StringBuilder();
            bool beforeCursor = true, nextToCursor = false;
            foreach (var node in elements)
            {
                var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(node.Value);
                string line = HttpUtility.HtmlDecode(htmlDocument.DocumentNode.InnerText);
                if (nextToCursor)
                {
                    sb.Append(line);
                    nextToCursor = false;
                }
                else
                {
                    if (beforeCursor)
                    {
                        beforeCursor = !HasAttributeValue(node, "selected", "all");
                        nextToCursor = !beforeCursor && string.IsNullOrEmpty(line);
                    }
                    if (!nextToCursor)
                    {
                        // non-cursor element
                        sb.Append('\n').Append(line);
                    }
                }
            }
            if (sb.Length > 0 && sb[0] == '\n')
            {
                sb.Remove(0, 1);
            }
            return sb.ToString();
        }

        public void EnumTextHtml(IEnumerable<XElement> elements, Action<HtmlAgilityPack.HtmlDocument> action)
        {
            void SetHtml(XElement element)
            {
                var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(element.Value);
                action.Invoke(htmlDocument);
                element.SetValue(htmlDocument.DocumentNode.InnerHtml);
            }

            foreach (var element in elements)
            {
                if (element.Name == doc.Root.Name.Namespace + "T")
                {
                    SetHtml(element);
                }
                else
                {
                    foreach (var tElement in Descendants(element, "T"))
                    {
                        SetHtml(element);
                    }
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
            return GetInnerHtmlText(Descendants(element, "T"));
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
                if (exist)
                {
                    element.SetAttributeValue(attrName, value);
                }
                else if (attr != null)
                {
                    attr.Remove();
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
            var elements = Descendants(doc.Root, "QuickStyleDef");
            if (elements != null)
            {
                foreach (var element in elements)
                {
                    element.SetAttributeValue(attrName, attrValue);
                }
            }
        }

        public void Save()
        {
            try
            {
                app.UpdatePageContent(doc.ToString(), DateTime.MinValue);
            }catch(Exception err)
            {
                Forms.MessageBox.Show("Error while update page : " + err.ToString());
            }
        }
    }
}
