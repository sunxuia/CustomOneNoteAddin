using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using System.Threading;
using System.Windows.Forms;
namespace OneNoteAddin
{
    /// <summary>
    /// demo api command of ribbon
    /// </summary>
    public partial class Ribbon : IDTExtensibility2, IRibbonExtensibility
    {
        private ShowForm showForm = new ShowForm();

        public void StartDebug(IRibbonControl control)
        {
            System.Diagnostics.Debugger.Launch();
        }

        public void ShowHierachy(IRibbonControl control)
        {
            string xml;
            app.GetHierarchy("", HierarchyScope.hsPages, out xml);
            showForm.Show(xml);
        }

        public void ShowCurrentWindow(IRibbonControl control)
        {
            showForm.Show($@"app.Windows.CurrentWindow
    CurrentNotebookId : {app.Windows.CurrentWindow.CurrentNotebookId}
CurrentSectionGroupId : {app.Windows.CurrentWindow.CurrentSectionGroupId}
     CurrentSectionId : {app.Windows.CurrentWindow.CurrentSectionId}
        CurrentPageId : {app.Windows.CurrentWindow.CurrentPageId}
               Active : {app.Windows.CurrentWindow.Active}
       DockedLocation : {app.Windows.CurrentWindow.DockedLocation}
         FullPageView : {app.Windows.CurrentWindow.FullPageView}
             SideNote : {app.Windows.CurrentWindow.SideNote}
");
        }

        public void ShowCurrentPageXml(IRibbonControl control)
        {
            string pageId = app.Windows.CurrentWindow.CurrentPageId;
            string xml;
            app.GetPageContent(pageId, out xml);
            showForm.Show(xml);
        }

        public void SetPageToAExamplePage(IRibbonControl control)
        {
            app.DeletePageContent(app.Windows.CurrentWindow.CurrentPageId, app.Windows.CurrentWindow.CurrentPageId);
            string xml = $@"<?xml version=""1.0""?>
<one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"" ID=""{app.Windows.CurrentWindow.CurrentPageId}"" name=""Sample Title"" pageLevel=""1"" isCurrentlyViewed=""true"">
    <one:QuickStyleDef index=""0"" name=""PageTitle"" fontColor=""automatic"" highlightColor=""automatic"" font=""Calibri"" fontSize=""20.0"" spaceBefore=""0.0"" spaceAfter=""0.0""/>
    <one:PageSettings RTL=""false"" color=""automatic"">
        <one:PageSize>
            <one:Automatic/>
        </one:PageSize>
        <one:RuleLines visible=""false""/>
    </one:PageSettings>
    <one:Title lang=""zh-CN"">
        <one:OE alignment=""left"" quickStyleIndex=""0"">
            <one:T><![CDATA[Sample Page]]></one:T>
        </one:OE>
    </one:Title>
</one:Page>
";
            app.UpdatePageContent(xml);
        }

        public void InsertATable(IRibbonControl control)
        {
            InsertWithClipboard(TextDataFormat.Html, @"Version:1.0
StartHTML:0000000105
EndHTML:0000001067
StartFragment:0000000527
EndFragment:0000001027

<html xmlns:o=""urn:schemas-microsoft-com:office:office""
xmlns:dt=""uuid:C2F41010-65B3-11d1-A29F-00AA00C14882""
xmlns=""http://www.w3.org/TR/REC-html40"">

<head>
<meta http-equiv=Content-Type content=""text/html; charset=utf-8"">
<meta name=ProgId content=OneNote.File>
<meta name=Generator content=""Microsoft OneNote 15"">
</head>

<body lang=zh-CN style='font-family:Calibri;font-size:11.0pt'>
<!--StartFragment-->

<div style='direction:ltr'>

<table border=1 cellpadding=0 cellspacing=0 valign=top style='direction:ltr;
 border-collapse:collapse;border-style:solid;border-color:#A3A3A3;border-width:
 1pt' title="""" summary="""">
 <tr>
  <td style='border-style:solid;border-color:#A3A3A3;border-width:1pt;
  background-color:#DEEBF6;vertical-align:top;width:.6381in;padding:4pt 4pt 4pt 4pt'>
  <p style='margin:0in;font-family:Calibri;font-size:11.0pt'>&nbsp;</p>
  </td>
 </tr>
</table>

</div>

<!--EndFragment-->
</body>

</html>");
        }

        private void InsertWithClipboard(TextDataFormat format, string context)
        {
            RunInThread(() =>
            {
                Clipboard.SetText(context, format);
            });
            SendKeys.SendWait("^(v)");
        }

        private void RunInThread(ThreadStart method)
        {
            Thread workProcess = new Thread(method);
            workProcess.SetApartmentState(ApartmentState.STA);//设置成单线程
            workProcess.Start();
            workProcess.Join();
        }

    }
}
