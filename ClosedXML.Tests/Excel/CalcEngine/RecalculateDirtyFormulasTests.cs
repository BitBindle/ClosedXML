using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// After a change of some cells, a caller needs the new results of the formulas that read the cells. A recalculation of
    /// all formulas does the work again for each formula. A read of <see cref="IXLCell.Value"/> calculates only the dirty
    /// formulas, but the caller must know the cells, and the first failure throws.
    /// </summary>
    [TestFixture]
    public class RecalculateDirtyFormulasTests
    {
        [Test]
        public void Recalculation_of_dirty_formulas_calculates_only_the_dependents_of_a_change()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").FormulaA1 = "A1*2";
            ws.Cell("C1").FormulaA1 = "B1+1";

            // RAND gives a new number for each calculation, so a calculation of D1 is visible.
            ws.Cell("D1").FormulaA1 = "RAND()";
            wb.RecalculateAllFormulas();
            var random = ws.Cell("D1").CachedValue;

            ws.Cell("A1").Value = 5;
            wb.RecalculateDirtyFormulas();

            Assert.AreEqual(10, ws.Cell("B1").CachedValue);
            Assert.AreEqual(11, ws.Cell("C1").CachedValue);
            Assert.AreEqual(random, ws.Cell("D1").CachedValue);
            Assert.IsFalse(ws.Cell("C1").NeedsRecalculation);
        }

        [Test]
        public void Recalculation_of_all_formulas_still_calculates_each_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "RAND()";
            wb.RecalculateAllFormulas();
            var random = ws.Cell("A1").CachedValue;

            wb.RecalculateAllFormulas();

            Assert.AreNotEqual(random, ws.Cell("A1").CachedValue);
        }

        [Test]
        public void Recalculation_of_dirty_formulas_calculates_new_formulas()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 3;
            ws.Cell("B1").FormulaA1 = "A1*2";

            wb.RecalculateDirtyFormulas();

            Assert.AreEqual(6, ws.Cell("B1").CachedValue);
        }

        [Test]
        public void Recalculation_of_dirty_formulas_collects_a_failed_cell_and_continues()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").FormulaA1 = "A1+1";
            ws.Cell("C1").FormulaA1 = "A1*10";
            wb.RecalculateAllFormulas();

            // The reference to a spill range is not implemented, so the calculation of B1 throws.
            ws.Cell("B1").FormulaA1 = "A1#";

            // A set of a formula moves the cell to the end of the calculation chain, so C1 comes after B1.
            ws.Cell("C1").FormulaA1 = "A1*10";
            ws.Cell("A1").Value = 2;
            Assert.Catch(() => wb.RecalculateDirtyFormulas());
            Assert.IsTrue(ws.Cell("C1").NeedsRecalculation);

            var failed = new List<IXLCell>();
            wb.RecalculateDirtyFormulas(failed);

            CollectionAssert.AreEqual(new[] { "S!B1" }, failed.Select(c => c.Address.ToString(XLReferenceStyle.A1, true)));
            Assert.AreEqual(20, ws.Cell("C1").CachedValue);
            Assert.IsFalse(ws.Cell("B1").NeedsRecalculation);
        }

        [Test]
        public void Recalculation_of_dirty_formulas_without_a_collection_throws_at_a_failed_cell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("S");
            ws.Cell("A1").FormulaA1 = "A2";
            ws.Cell("A2").FormulaA1 = "A1";

            Assert.That(() => wb.RecalculateDirtyFormulas(),
                Throws.InvalidOperationException.With.Message.EqualTo("Formula in a cell S!A1 is part of a cycle."));
        }

        [Test]
        public void Recalculation_of_dirty_formulas_requires_a_collection()
        {
            using var wb = new XLWorkbook();

            Assert.Throws<ArgumentNullException>(() => wb.RecalculateDirtyFormulas(null!));
        }

        [Test]
        public void Mark_of_dependents_makes_a_formula_that_reads_a_hidden_row_dirty()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("A3").Value = 4;
            ws.Cell("B1").FormulaA1 = "SUBTOTAL(109,A1:A3)";
            ws.Cell("C1").FormulaA1 = "B1*10";
            ws.Cell("D1").FormulaA1 = "SUM(A5:A6)";
            wb.RecalculateAllFormulas();

            // SUBTOTAL 109 ignores hidden rows, but a hide of a row does not change the value of a cell.
            ws.Row(2).Hide();
            Assert.IsFalse(ws.Cell("B1").NeedsRecalculation);

            wb.MarkDependentsDirty(ws.Row(2));

            Assert.IsTrue(ws.Cell("B1").NeedsRecalculation);
            Assert.IsTrue(ws.Cell("C1").NeedsRecalculation);
            Assert.IsFalse(ws.Cell("D1").NeedsRecalculation);
            wb.RecalculateDirtyFormulas();
            Assert.AreEqual(5, ws.Cell("B1").CachedValue);
            Assert.AreEqual(50, ws.Cell("C1").CachedValue);
        }

        [Test]
        public void Mark_of_dependents_accepts_a_range_on_another_sheet()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet("Data");
            var report = wb.AddWorksheet("Report");
            report.Cell("A1").FormulaA1 = "SUM(Data!B2:C3)";
            report.Cell("A2").FormulaA1 = "SUM(Data!D2:D3)";
            wb.RecalculateAllFormulas();

            wb.MarkDependentsDirty(data.Range("C3:D4"));

            Assert.IsTrue(report.Cell("A1").NeedsRecalculation);
            Assert.IsTrue(report.Cell("A2").NeedsRecalculation);
        }

        [Test]
        public void Mark_of_dependents_does_not_mark_the_formula_in_the_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "1+1";
            wb.RecalculateAllFormulas();

            wb.MarkDependentsDirty(ws.Cell("A1").AsRange());

            Assert.IsFalse(ws.Cell("A1").NeedsRecalculation);
        }

        [Test]
        public void Mark_of_dependents_requires_a_range_of_the_workbook()
        {
            using var wb = new XLWorkbook();
            using var other = new XLWorkbook();
            var range = other.AddWorksheet().Range("A1:B2");

            Assert.Throws<ArgumentNullException>(() => wb.MarkDependentsDirty(null!));
            Assert.Throws<ArgumentException>(() => wb.MarkDependentsDirty(range));

            var ws = wb.AddWorksheet();
            var deletedRange = ws.Range("A2:B2");
            ws.Row(2).Delete();
            Assert.Throws<ArgumentException>(() => wb.MarkDependentsDirty(deletedRange));
        }
    }
}
