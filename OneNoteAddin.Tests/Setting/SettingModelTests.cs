using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteAddin.Setting;

namespace OneNoteAddin.Tests
{
    [TestClass]
    public class SettingModelTests
    {
        [TestMethod]
        public void TestSettingModel()
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)), "setting.xml");
            string xml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<setting>
  <codeStyles>
    <codeStyle>None</codeStyle>
  </codeStyles>
  <defaultValues>
    <defaultValue id=""cmbFont1"" value=""Microsoft YaHei Mono""></defaultValue>
  </defaultValues>
  <tables>
    <table label=""Blue Title Table"" size=""large"" row=""2"" column=""2"" foreColor=""ffffff"" backColor=""2e75b5"" headInLeft=""true"" />
  </tables>
</setting>
";
            File.WriteAllText(path, xml, Encoding.UTF8);

            SettingModel setting = new SettingModel(path);

            Assert.AreEqual(path, setting.FilePath);

            Assert.AreEqual(1, setting.CodeStyles.Count);
            Assert.AreEqual("None", setting.CodeStyles[0]);

            Assert.AreEqual(1, setting.DefaultValues.Count);
            Assert.AreEqual("cmbFont1", setting.DefaultValues[0].Id);
            Assert.AreEqual("Microsoft YaHei Mono", setting.DefaultValues[0].Value);

            Assert.AreEqual(1, setting.Tables.Count);
            Assert.AreEqual("Blue Title Table", setting.Tables[0].Label);
            Assert.AreEqual("large", setting.Tables[0].Size);
            Assert.AreEqual(2, setting.Tables[0].Row);
            Assert.AreEqual(2, setting.Tables[0].Column);
            Assert.AreEqual("ffffff", setting.Tables[0].ForeColor);
            Assert.AreEqual("2e75b5", setting.Tables[0].BackColor);
            Assert.AreEqual(true, setting.Tables[0].HeadInLeft);
        }

        [TestMethod]
        public void TestWriteXml()
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8)), "setting.xml");
            string xml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<setting>
  <codeStyles />
  <defaultValues />
  <tables />
</setting>
";
            File.WriteAllText(path, xml, Encoding.UTF8);
            SettingModel setting = new SettingModel(path);

            setting.CodeStyles.Add("None");
            setting.DefaultValues.Add(new DefaultValueModel()
            {
                Id = "cmbFont1",
                Value = "Microsoft YaHei Mono"
            });
            setting.Tables.Add(new TableModel()
            {
                Label = "Blue Title Table",
                Size = "large",
                Row = 2,
                Column = 2,
                ForeColor = "ffffff",
                BackColor = "2e75b5",
                HeadInLeft = true
            });
            setting.WriteXml();

            SettingModel saved = new SettingModel(path);
            Assert.AreEqual(setting.CodeStyles.Count, saved.CodeStyles.Count);
            Assert.AreEqual(setting.CodeStyles[0], saved.CodeStyles[0]);

            Assert.AreEqual(setting.DefaultValues.Count, saved.DefaultValues.Count);
            Assert.AreEqual(setting.DefaultValues[0].Id, saved.DefaultValues[0].Id);
            Assert.AreEqual(setting.DefaultValues[0].Value, saved.DefaultValues[0].Value);

            Assert.AreEqual(setting.Tables.Count, saved.Tables.Count);
            Assert.AreEqual(setting.Tables[0].Label, saved.Tables[0].Label);
            Assert.AreEqual(setting.Tables[0].Size, saved.Tables[0].Size);
            Assert.AreEqual(setting.Tables[0].Row, saved.Tables[0].Row);
            Assert.AreEqual(setting.Tables[0].Column, saved.Tables[0].Column);
            Assert.AreEqual(setting.Tables[0].ForeColor, saved.Tables[0].ForeColor);
            Assert.AreEqual(setting.Tables[0].BackColor, saved.Tables[0].BackColor);
            Assert.AreEqual(setting.Tables[0].HeadInLeft, saved.Tables[0].HeadInLeft);
        }
    }
}
