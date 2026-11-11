using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class CalcEngineExceptionTests
    {
        [OneTimeSetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        [Test]
        public void InvalidCharNumber()
        {
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("CHAR(-2)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("CHAR(270)"));
        }

        [Test]
        public void DivisionByZero()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("0/0"));
            Assert.AreEqual(XLError.DivisionByZero, new XLWorkbook().AddWorksheet().Evaluate("0/0"));
        }

        [Test]
        public void InvalidFunction()
        {
            Assert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("XXX(A1:A2)"));

            var ws = new XLWorkbook().AddWorksheet();
            Assert.AreEqual(XLError.NameNotRecognized, ws.Evaluate("XXX(A1:A2)"));
        }

        [Test]
        public void NestedNameNotRecognizedException()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").SetFormulaA1("=XXX");
            ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

            Assert.AreEqual("Success", ws.Cell("A2").Value);
        }

        [Test]
        public void Recalculation_with_failed_cells_continues_after_cycle()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").FormulaA1 = "A2";
            ws.Cell("A2").FormulaA1 = "A1";
            ws.Cell("B1").FormulaA1 = "1+1";
            Assert.Throws<InvalidOperationException>(() => wb.RecalculateAllFormulas());

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            // The cycle is detected at A1. A2 uses the kept value of A1.
            CollectionAssert.AreEqual(new[] { "S!A1" }, failed.Select(c => c.Address.ToString(XLReferenceStyle.A1, true)));
            Assert.AreEqual(0, ws.Cell("A2").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
        }

        [TestCase("1+")] // Doesn't parse
        [TestCase("A2#")] // Throws NotImplementedException
        public void Recalculation_with_failed_cells_continues_after_failed_formula(string formula)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").FormulaA1 = formula;
            ws.Cell("B1").FormulaA1 = "A1+1";
            ws.Cell("C1").FormulaA1 = "1+1";
            Assert.Catch(() => wb.RecalculateAllFormulas());

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            CollectionAssert.AreEqual(new[] { "S!A1" }, failed.Select(c => c.Address.ToString(XLReferenceStyle.A1, true)));
            Assert.AreEqual(Blank.Value, ws.Cell("A1").CachedValue);
            Assert.AreEqual(1, ws.Cell("B1").Value); // Calculated with the kept value of A1
            Assert.AreEqual(2, ws.Cell("C1").Value);
        }

        [Test]
        public void Recalculation_with_failed_cells_throws_on_cancellation()
        {
            var cts = new CancellationTokenSource();
            using var wb = new XLWorkbook(new LoadOptions { CancellationToken = cts.Token });
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "REPT(\"A\",10)";

            cts.Cancel();
            Assert.Throws<OperationCanceledException>(() => wb.RecalculateAllFormulas(new List<IXLCell>()));
        }
    }
}
