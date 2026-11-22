using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Cells
{
    /// <summary>
    /// Tests for the references in cell formulas after an insert or a delete of cells.
    /// </summary>
    [TestFixture]
    public class FormulaReferenceShiftTests
    {
        [TestCase("LOG10(A5)", "LOG10(A7)")]
        [TestCase("ATAN2(A5,B5)", "ATAN2(A7,B7)")]
        [TestCase("\"A5\"&A5", "\"A5\"&A7")]
        [TestCase("\"x\"\"A5\"&A5", "\"x\"\"A5\"&A7")]
        [TestCase("[1]Sheet1!A5", "[1]Sheet1!A5")]
        [TestCase("SUM(2:10)", "SUM(2:12)")]
        [TestCase("SUM(1:1048576)", "SUM(1:1048576)")]
        [TestCase("SUM(A1:B2)", "SUM(A1:B2)")]
        [TestCase("Sheet1!$A$5+'Sheet1'!A5", "Sheet1!$A$7+Sheet1!A7")]
        [TestCase("Other!A5", "Other!A5")]
        [TestCase("A1048575#", "#REF!")]
        [TestCase("-A5%", "-A7%")]
        [Issue("75")]
        public void Insert_rows_shifts_only_the_references_below_the_insert(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            wb.AddWorksheet("Other");
            ws.Cell("H1").FormulaA1 = formula;

            ws.Row(3).InsertRowsAbove(2);

            Assert.AreEqual(expected, ws.Cell("H1").FormulaA1);
        }

        [TestCase("SUM(B:D)", "SUM(B:E)")]
        [TestCase("SUM($C:$C)", "SUM($D:$D)")]
        [TestCase("SUM(A:B)", "SUM(A:B)")]
        [TestCase("\"C1\"&C1", "\"C1\"&D1")]
        [Issue("75")]
        public void Insert_columns_shifts_only_the_references_right_of_the_insert(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A10").FormulaA1 = formula;

            ws.Column("C").InsertColumnsBefore(1);

            Assert.AreEqual(expected, ws.Cell("A10").FormulaA1);
        }

        [TestCase("A3", "#REF!")]
        [TestCase("A3#", "#REF!")]
        [TestCase("A4#", "A3#")]
        [TestCase("A4+A5", "A3+A4")]
        [TestCase("SUM($3:$5)", "SUM($3:$4)")]
        [TestCase("SUM(A1:A10)", "SUM(A1:A9)")]
        [TestCase("SUM(A1:A2)", "SUM(A1:A2)")]
        [Issue("75")]
        public void Delete_rows_shifts_the_references_and_removes_deleted_cells(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("H1").FormulaA1 = formula;

            ws.Row(3).Delete();

            Assert.AreEqual(expected, ws.Cell("H1").FormulaA1);
        }

        [TestCase("C1", "#REF!")]
        [TestCase("D1", "C1")]
        [TestCase("SUM(B1:E1)", "SUM(B1:D1)")]
        [Issue("75")]
        public void Delete_columns_shifts_the_references_and_removes_deleted_cells(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A10").FormulaA1 = formula;

            ws.Column("C").Delete();

            Assert.AreEqual(expected, ws.Cell("A10").FormulaA1);
        }

        [Test]
        [Issue("75")]
        public void Insert_into_a_range_shifts_only_the_references_in_the_columns_of_the_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("H1").FormulaA1 = "C5+F5";

            ws.Range("B3:D3").InsertRowsAbove(1);

            Assert.AreEqual("C6+F5", ws.Cell("H1").FormulaA1);
        }

        [Test]
        [Issue("75")]
        public void Formula_on_other_sheet_is_shifted()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var other = wb.AddWorksheet("Other");
            other.Cell("A1").FormulaA1 = "Sheet1!A5+A5";

            ws.Row(3).InsertRowsAbove(1);

            Assert.AreEqual("Sheet1!A6+A5", other.Cell("A1").FormulaA1);
        }

        [Test]
        [Issue("75")]
        public void Unparsable_formula_stays_the_same()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("H1").FormulaA1 = "SUM(A5";

            ws.Row(3).InsertRowsAbove(1);

            Assert.AreEqual("SUM(A5", ws.Cell("H1").FormulaA1);
        }

        [Test]
        [Issue("75")]
        public void Formula_without_a_shifted_reference_gets_a_new_value()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("B10").Value = 5;
            ws.Cell("C10").Value = 1;
            ws.Cell("H1").FormulaA1 = "SUM(A1:C10)";
            Assert.AreEqual(6, ws.Cell("H1").Value);

            // The insert moves B10 out of the area, but the reference does not change.
            ws.Range("B3:B3").InsertRowsAbove(1);

            Assert.AreEqual("SUM(A1:C10)", ws.Cell("H1").FormulaA1);
            Assert.AreEqual(1, ws.Cell("H1").Value);
        }
    }
}
