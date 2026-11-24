#nullable enable

using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// A caller that evaluates a workbook from an unknown source can block the defined names and
    /// the tables that it doesn't trust or doesn't load. A formula that uses a blocked name fails.
    /// </summary>
    [TestFixture]
    public class CalculationNameFilterTests
    {
        [Test]
        public void Blocked_workbook_name_fails_the_cell()
        {
            using var wb = Workbook((sheet, name) => !(sheet is null && name == "Rate"));
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Rate", "S1!$A$1");
            wb.DefinedNames.Add("Other", "S1!$A$3");
            ws.Cell("B1").FormulaA1 = "Rate*10";
            ws.Cell("B2").FormulaA1 = "Other*10";
            ws.Cell("B3").FormulaA1 = "B1+1";

            Assert.That(() => ws.Cell("B1").Value, Throws.InvalidOperationException);

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed, Is.EqualTo(new[] { ws.Cell("B1") }));
            Assert.That(ws.Cell("B2").Value, Is.EqualTo(XLCellValue.FromObject(30)));
            Assert.That(ws.Cell("B3").Value, Is.EqualTo(XLCellValue.FromObject(1))); // The kept value of B1 is blank.
        }

        [TestCase("S1", "Rate", true)] // The blocked sheet name hides the workbook name.
        [TestCase("S1", "S2!Rate", false)]
        [TestCase("S2", "Rate", false)]
        [TestCase("S2", "S1!Rate", true)]
        public void Blocked_sheet_name_hides_the_workbook_name_on_its_sheet(string formulaSheet, string formula, bool fails)
        {
            using var wb = Workbook((sheet, _) => sheet != "S1");
            wb.AddWorksheet("S2");
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Rate", "S1!$A$1");
            ws.DefinedNames.Add("Rate", "S1!$A$2");
            var cell = wb.Worksheet(formulaSheet).Cell("D1");
            cell.FormulaA1 = formula;

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed, fails ? Is.EqualTo(new[] { cell }) : Is.Empty);
            if (!fails)
                Assert.That(cell.Value, Is.EqualTo(XLCellValue.FromObject(1)));
        }

        [Test]
        public void Sheet_name_hides_blocked_workbook_name_on_its_sheet()
        {
            using var wb = Workbook((sheet, _) => sheet is not null);
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Rate", "S1!$A$1");
            ws.DefinedNames.Add("Rate", "S1!$A$2");
            ws.Cell("B1").FormulaA1 = "Rate";

            Assert.That(ws.Cell("B1").Value, Is.EqualTo(XLCellValue.FromObject(2)));
        }

        [Test]
        public void Allow_list_with_workbook_names_allows_them_on_each_sheet()
        {
            var trusted = new HashSet<(string?, string)> { (null, "Rate") };
            using var wb = Workbook((sheet, name) => trusted.Contains((sheet, name)));
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Rate", "S1!$A$1");
            wb.DefinedNames.Add("Other", "S1!$A$1");
            ws.Cell("B1").FormulaA1 = "Rate*2";
            ws.Cell("B2").FormulaA1 = "rate*2"; // The filter gets the stored name.
            ws.Cell("B3").FormulaA1 = "Other*2";

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed, Is.EqualTo(new[] { ws.Cell("B3") }));
            Assert.That(ws.Cell("B1").Value, Is.EqualTo(XLCellValue.FromObject(2)));
            Assert.That(ws.Cell("B2").Value, Is.EqualTo(XLCellValue.FromObject(2)));
        }

        [TestCase("Missing*2", true)]
        [TestCase("Unknown*2", false)]
        public void Blocked_name_that_workbook_does_not_have_fails_the_cell(string formula, bool fails)
        {
            using var wb = Workbook((_, name) => name != "Missing");
            var ws = wb.Worksheet("S1");
            ws.Cell("B1").FormulaA1 = formula;

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed.Count, Is.EqualTo(fails ? 1 : 0));
            if (!fails)
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(XLCellValue.FromObject(XLError.NameNotRecognized)));
        }

        [TestCase("INDIRECT(\"Rate\")", true)]
        [TestCase("INDIRECT(\"RATE\")", true)] // The filter gets the stored name, not the text.
        [TestCase("rate*1", true)]
        [TestCase("INDIRECT(\"Other\")", false)]
        public void Indirect_of_blocked_name_fails_the_cell(string formula, bool fails)
        {
            using var wb = Workbook((_, name) => name != "Rate");
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Rate", "S1!$A$1");
            wb.DefinedNames.Add("Other", "S1!$A$1");
            ws.Cell("B1").FormulaA1 = formula;

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed.Count, Is.EqualTo(fails ? 1 : 0));
            if (!fails)
                Assert.That(ws.Cell("B1").Value, Is.EqualTo(XLCellValue.FromObject(1)));
        }

        [Test]
        public void Dependency_tree_does_not_read_formula_of_blocked_name()
        {
            // A formula that reads the cell of an allowed name becomes dirty when the cell changes.
            // The dependency tree doesn't read the formula of a blocked name, so its formula doesn't.
            using var wb = Workbook((_, name) => name != "Blocked");
            var ws = wb.Worksheet("S1");
            wb.DefinedNames.Add("Blocked", "S1!$A$1");
            wb.DefinedNames.Add("Allowed", "S1!$A$1");
            ws.Cell("B1").FormulaA1 = "Blocked";
            ws.Cell("B2").FormulaA1 = "Allowed";
            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);
            Assert.That(failed, Is.EqualTo(new[] { ws.Cell("B1") }));

            ws.Cell("A1").Value = 5;

            Assert.That(ws.Cell("B2").NeedsRecalculation, Is.True);
            Assert.That(ws.Cell("B1").NeedsRecalculation, Is.False);
        }

        [TestCase("SUM(Pastries[Price])")]
        [TestCase("SUM(INDIRECT(\"Pastries[Price]\"))")]
        [TestCase("SUM(pastries[Price])")] // The filter gets the stored name, not the text.
        [TestCase("SUM(INDIRECT(\"PASTRIES[Price]\"))")]
        [TestCase("SUM(Missing[Price])")] // The workbook doesn't have the table.
        public void Blocked_table_fails_the_cell(string formula)
        {
            using var wb = Workbook((_, name) => name is not ("Pastries" or "Missing"));
            var ws = wb.Worksheet("S1");
            AddTable(ws, "Pastries");
            AddTable(ws, "Cakes", "E3");
            ws.Cell("A1").FormulaA1 = formula;
            ws.Cell("A2").FormulaA1 = "SUM(Cakes[Price])";

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed, Is.EqualTo(new[] { ws.Cell("A1") }));
            Assert.That(ws.Cell("A2").Value, Is.EqualTo(XLCellValue.FromObject(16)));
        }

        [Test]
        public void Blocked_table_fails_a_reference_without_table_name()
        {
            using var wb = Workbook((_, name) => name != "Pastries");
            var ws = wb.Worksheet("S1");
            AddTable(ws, "Pastries");

            Assert.That(() => ws.Evaluate("SUM([Price])", "C4"), Throws.InvalidOperationException);
        }

        [Test]
        public void Default_options_evaluate_each_name_and_table()
        {
            Assert.That(new LoadOptions().CalculationNameFilter, Is.Null);
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S1");
            ws.Cell("A1").Value = 1;
            wb.DefinedNames.Add("Rate", "S1!$A$1");
            ws.DefinedNames.Add("Local", "S1!$A$1");
            AddTable(ws, "Pastries");
            ws.Cell("B1").FormulaA1 = "Rate+Local+INDIRECT(\"Rate\")";
            ws.Cell("B2").FormulaA1 = "SUM(Pastries[Price])+SUM(INDIRECT(\"Pastries[Price]\"))";
            ws.Cell("B3").FormulaA1 = "SUM(Missing[Price])";

            var failed = new List<IXLCell>();
            wb.RecalculateAllFormulas(failed);

            Assert.That(failed, Is.Empty);
            Assert.That(ws.Cell("B1").Value, Is.EqualTo(XLCellValue.FromObject(3)));
            Assert.That(ws.Cell("B2").Value, Is.EqualTo(XLCellValue.FromObject(32)));
            Assert.That(ws.Cell("B3").Value, Is.EqualTo(XLCellValue.FromObject(XLError.CellReference)));
        }

        private static XLWorkbook Workbook(Func<string?, string, bool> filter)
        {
            var wb = new XLWorkbook(new LoadOptions { CalculationNameFilter = filter });
            var ws = wb.AddWorksheet("S1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 3;
            return wb;
        }

        private static void AddTable(IXLWorksheet ws, string name, string origin = "B3")
        {
            ws.Cell(origin).InsertTable(
                [
                    new { Name = "Cake", Price = 7 },
                    new { Name = "Pie", Price = 9 }
                ],
                name);
        }
    }
}
