using ClosedXML.Excel;
using ClosedXML.Utils;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Saving
{
    /// <summary>
    /// The file format defines some text attributes as <c>ST_Xstring</c>. A save must encode them and a load must
    /// decode them, so the text stays the same.
    /// </summary>
    [TestFixture]
    public class TextAttributeTests
    {
        [TestCase("a\rb", false)]
        [TestCase("a_x0041_b", false)]
        [TestCase("a\u0001b", false, TestName = "{m}(U+0001,False)")]
        [TestCase("a\rb", true)]
        [TestCase("a_x0041_b", true)]
        [TestCase("a\u0001b", true, TestName = "{m}(U+0001,True)")]
        [Issue("107")]
        public void Save_and_load_keep_the_messages_of_a_data_validation(string text, bool otherSheet)
        {
            // A data validation with a reference to another sheet goes into the x14 extension.
            using var stream = SaveAndCheck(wb =>
            {
                var ws = wb.AddWorksheet("Data");
                wb.AddWorksheet("Other");
                var dv = ws.Range("A1:A2").CreateDataValidation();
                dv.List(otherSheet ? "Other!A1:A3" : "B1:B3");
                dv.InputTitle = text + "1";
                dv.InputMessage = text + "2";
                dv.ErrorTitle = text + "3";
                dv.ErrorMessage = text + "4";
            }, document =>
            {
                var dataValidation = document.WorkbookPart!.WorksheetParts
                    .SelectMany(p => p.Worksheet.Descendants())
                    .Single(e => e.LocalName == "dataValidation");
                Assert.AreEqual(XmlEncoder.EncodeString(text + "1"), dataValidation.GetAttribute("promptTitle", "").Value);
                Assert.AreEqual(XmlEncoder.EncodeString(text + "2"), dataValidation.GetAttribute("prompt", "").Value);
                Assert.AreEqual(XmlEncoder.EncodeString(text + "3"), dataValidation.GetAttribute("errorTitle", "").Value);
                Assert.AreEqual(XmlEncoder.EncodeString(text + "4"), dataValidation.GetAttribute("error", "").Value);
            });

            using var wb = new XLWorkbook(stream);
            var loaded = wb.Worksheet("Data").DataValidations.Single();
            Assert.AreEqual(text + "1", loaded.InputTitle);
            Assert.AreEqual(text + "2", loaded.InputMessage);
            Assert.AreEqual(text + "3", loaded.ErrorTitle);
            Assert.AreEqual(text + "4", loaded.ErrorMessage);
        }

        [TestCase("a\rb")]
        [TestCase("a_x0041_b")]
        [TestCase("a\u0001b", TestName = "{m}(U+0001)")]
        [Issue("107")]
        public void Save_and_load_keep_the_text_of_a_conditional_format(string text)
        {
            using var stream = SaveAndCheck(wb =>
            {
                var ws = wb.AddWorksheet();
                ws.Range("A1:A2").AddConditionalFormat().WhenContains(text).Fill.SetBackgroundColor(XLColor.Red);
            }, document =>
            {
                var rule = document.WorkbookPart!.WorksheetParts.Single().Worksheet
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule>().Single();
                Assert.AreEqual(XmlEncoder.EncodeString(text), rule.Text!.Value);
            });

            using var wb = new XLWorkbook(stream);
            var conditionalFormat = wb.Worksheets.Single().ConditionalFormats.Single();
            Assert.AreEqual(XLConditionalFormatType.ContainsText, conditionalFormat.ConditionalFormatType);
            Assert.AreEqual(text, conditionalFormat.Values[1].Value);
        }

        private static MemoryStream SaveAndCheck(Action<XLWorkbook> create, Action<SpreadsheetDocument> check)
        {
            var stream = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                create(wb);
                wb.SaveAs(stream);
            }

            stream.Position = 0;
            using (var document = SpreadsheetDocument.Open(stream, false))
                check(document);

            stream.Position = 0;
            return stream;
        }
    }
}
