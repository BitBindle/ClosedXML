#nullable enable

using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// Excel doesn't accept a formula that calls a function with a wrong number of arguments, but
    /// a file from another application can contain one. LibreOffice gives <c>#VALUE!</c> for the
    /// whole formula. <see cref="LoadOptions.WrongArgumentCountAsErrorValue"/> gives the same
    /// result. Without the option, the parse of the formula throws.
    /// </summary>
    [TestFixture]
    public class WrongArgumentCountTests
    {
        [TestCase("YEAR()")]
        [TestCase("NOW(1)")]
        [TestCase("ABS(1,2)")]
        [TestCase("ROUND(1,2,3)")]
        [TestCase("IF()")]
        [TestCase("DATE(2020,1)")]
        [TestCase("_xlfn.DAYS(1)")]
        [TestCase("1+YEAR()")]
        [TestCase("IFERROR(YEAR(),1)")] // The whole formula is the error, as in LibreOffice.
        [TestCase("ISERROR(NOW(1))")]
        [TestCase("IF(TRUE,1,YEAR())")]
        public void Formula_with_wrong_argument_count_is_value_error(string formula)
        {
            using var wb = Workbook();
            var ws = wb.Worksheet("S1");
            ws.Cell("A1").FormulaA1 = formula;

            Assert.That(ws.Cell("A1").Value, Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
        }

        [Test]
        public void Other_formulas_are_calculated()
        {
            using var wb = Workbook();
            var ws = wb.Worksheet("S1");
            ws.Cell("A1").FormulaA1 = "YEAR()";
            ws.Cell("A2").FormulaA1 = "A1";
            ws.Cell("A3").FormulaA1 = "IFERROR(A1,5)";
            ws.Cell("A4").FormulaA1 = "YEAR(1)";

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed, Is.Empty);
            Assert.That(ws.Cell("A2").Value, Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
            Assert.That(ws.Cell("A3").Value, Is.EqualTo(XLCellValue.FromObject(5)));
            Assert.That(ws.Cell("A4").Value, Is.EqualTo(XLCellValue.FromObject(1900)));
        }

        [Test]
        public void Evaluate_and_defined_name_give_value_error()
        {
            using var wb = Workbook();
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Wrong", "YEAR()");
            ws.Cell("A1").FormulaA1 = "Wrong";

            Assert.That(ws.Evaluate("YEAR()"), Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
            Assert.That(wb.Evaluate("NOW(1)"), Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
            Assert.That(ws.Cell("A1").Value, Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
        }

        [Test]
        public void Load_with_recalculation_gives_value_error()
        {
            using var stream = new MemoryStream();
            using (var source = new XLWorkbook())
            {
                var sheet = source.AddWorksheet("S1");
                sheet.Cell("A1").FormulaA1 = "YEAR()";
                sheet.Cell("A2").FormulaA1 = "1+1";
                source.SaveAs(stream);
            }

            stream.Position = 0;
            using var wb = new XLWorkbook(stream, new LoadOptions { WrongArgumentCountAsErrorValue = true, RecalculateAllFormulas = true });
            var ws = wb.Worksheet("S1");

            Assert.That(ws.Cell("A1").CachedValue, Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
            Assert.That(ws.Cell("A2").CachedValue, Is.EqualTo(XLCellValue.FromObject(2)));
        }

        [Test]
        public void Unknown_function_and_syntax_error_do_not_change()
        {
            using var wb = Workbook();
            var ws = wb.Worksheet("S1");
            ws.Cell("A1").FormulaA1 = "NOSUCHFUNCTION()";

            Assert.That(ws.Cell("A1").Value, Is.EqualTo(XLCellValue.FromObject(XLError.NameNotRecognized)));
            Assert.That(() => ws.Evaluate("SUM(1"), Throws.TypeOf<ExpressionParseException>());
        }

        [TestCase("YEAR()+")]
        [TestCase("SUM(1,YEAR()")]
        [TestCase("YEAR()+SUM(1")]
        public void Syntax_error_after_wrong_argument_count_is_parse_error(string formula)
        {
            using var wb = Workbook();
            var ws = wb.Worksheet("S1");

            Assert.That(() => ws.Evaluate(formula), Throws.TypeOf<ExpressionParseException>());
        }

        [Test]
        public void Wrong_argument_count_in_unknown_function_is_value_error()
        {
            using var wb = Workbook();
            var ws = wb.Worksheet("S1");

            Assert.That(ws.Evaluate("NOSUCHFUNCTION(YEAR())"), Is.EqualTo(XLCellValue.FromObject(XLError.IncompatibleValue)));
        }

        [Test]
        public void Default_options_throw_for_wrong_argument_count()
        {
            Assert.That(new LoadOptions().WrongArgumentCountAsErrorValue, Is.False);
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S1");
            ws.Cell("A1").FormulaA1 = "YEAR()";
            ws.Cell("A2").FormulaA1 = "1+1";

            Assert.That(() => ws.Evaluate("NOW(1)"), Throws.TypeOf<ExpressionParseException>());
            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);
            Assert.That(failed, Is.EqualTo(new[] { ws.Cell("A1") }));
            Assert.That(ws.Cell("A2").Value, Is.EqualTo(XLCellValue.FromObject(2)));
        }

        private static XLWorkbook Workbook()
        {
            var wb = new XLWorkbook(new LoadOptions { WrongArgumentCountAsErrorValue = true });
            wb.AddWorksheet("S1");
            return wb;
        }
    }
}
