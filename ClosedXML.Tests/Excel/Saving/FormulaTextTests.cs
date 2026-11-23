using ClosedXML.Excel;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel.Saving
{
    /// <summary>
    /// Tests for formulas outside of cells. These formulas are ST_Formula values, and an ST_Formula is an
    /// ST_Xstring. The save must encode special characters, and the load must decode them.
    /// </summary>
    [TestFixture]
    public class FormulaTextTests
    {
        [TestCase("a\rb")]
        [TestCase("a_x0041_b")]
        [Issue("104")]
        public void Save_and_load_keep_the_formula_of_a_defined_name(string text)
        {
            var formula = Quote(text);

            using var ms = SaveAndCheck(
                wb =>
                {
                    wb.DefinedNames.Add("Workbook_name", formula);
                    wb.Worksheet(1).DefinedNames.Add("Sheet_name", formula);
                },
                wb =>
                {
                    Assert.AreEqual(formula, wb.DefinedName("Workbook_name")!.RefersTo);
                    Assert.AreEqual(formula, wb.Worksheet(1).DefinedName("Sheet_name").RefersTo);
                });
        }

        [TestCase("a\rb")]
        [TestCase("a_x0041_b")]
        [TestCase("a\u0001b", TestName = "{m}(U+0001)")]
        [Issue("104")]
        public void Save_and_load_keep_the_formula_of_a_data_validation(string text)
        {
            var formula = Quote(text);

            using var ms = SaveAndCheck(
                wb =>
                {
                    var ws = wb.Worksheet(1);
                    ws.Range("A1").CreateDataValidation().Custom(formula);
                    ws.Range("A2").CreateDataValidation().TextLength.Between(formula, "LEN(" + formula + ")");
                },
                wb =>
                {
                    var validations = wb.Worksheet(1).DataValidations.ToList();
                    Assert.AreEqual(2, validations.Count);
                    Assert.AreEqual(formula, validations[0].MinValue);
                    Assert.AreEqual(formula, validations[1].MinValue);
                    Assert.AreEqual("LEN(" + formula + ")", validations[1].MaxValue);
                });
        }

        [TestCase("a\rb")]
        [TestCase("a_x0041_b")]
        [Issue("104")]
        public void Save_and_load_keep_the_formulas_of_a_conditional_format(string text)
        {
            var formula = Quote(text);

            using var ms = SaveAndCheck(
                wb =>
                {
                    var ws = wb.Worksheet(1);
                    ws.Range("A1").AddConditionalFormat().WhenIsTrue(formula).Fill.SetBackgroundColor(XLColor.Red);
                    ws.Range("A2").AddConditionalFormat().WhenEquals(formula).Fill.SetBackgroundColor(XLColor.Red);
                },
                wb =>
                {
                    var formats = wb.Worksheet(1).ConditionalFormats.ToList();
                    Assert.AreEqual(2, formats.Count);
                    Assert.True(formats.All(f => f.Values[1]!.Value == formula));
                });
        }

        [TestCase("a\rb")]
        [TestCase("a_x0041_b")]
        [Issue("104")]
        public void Save_and_load_keep_the_formulas_of_a_table_column(string text)
        {
            var formula = Quote(text);

            using var ms = SaveAndCheck(
                wb =>
                {
                    var ws = wb.Worksheet(1);
                    ws.Cell("A1").Value = "Name";
                    ws.Cell("B1").Value = "Calculated";
                    ws.Cell("A2").Value = "x";
                    ws.Cell("A3").Value = "y";
                    var table = ws.Range("A1:B3").CreateTable();
                    table.ShowTotalsRow = true;
                    ws.Range("B2:B3").FormulaA1 = formula;
                    table.Field("Name").TotalsRowFormulaA1 = formula;
                },
                wb => Assert.AreEqual(formula, wb.Worksheet(1).Table(0).Field("Name").TotalsRowFormulaA1));

            using var document = SpreadsheetDocument.Open(ms, false);
            var tablePart = document.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single();
            var columns = tablePart.Table.TableColumns!.Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>().ToList();
            var expected = Quote(XmlEncoder.EncodeString(text));
            Assert.AreEqual(expected, columns[0].TotalsRowFormula!.Text);
            Assert.AreEqual(expected, columns[1].CalculatedColumnFormula!.Text);
        }

        private static string Quote(string text) => "\"" + text + "\"";

        private static MemoryStream SaveAndCheck(Action<XLWorkbook> create, Action<XLWorkbook> check)
        {
            var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet("Sheet1");
                create(wb);
                wb.SaveAs(ms);
            }

            ms.Position = 0;
            using (var wb = new XLWorkbook(ms))
                check(wb);

            ms.Position = 0;
            return ms;
        }
    }
}
