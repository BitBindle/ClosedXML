using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class ArrayFormulaTests
    {
        [Test]
        public void ArrayFormulaIsSaved()
        {
            TestHelper.CreateAndCompare(wb =>
            {
                var ws = wb.AddWorksheet();
                ws.Range("A1:B2").FormulaArrayA1 = "1+2";
            }, @"Other\Formulas\ArrayFormula.xlsx");
        }

        [Test]
        public void ArrayFormulaCanBeLoaded()
        {
            TestHelper.LoadAndAssert(wb =>
            {
                var ws = wb.Worksheets.First();

                foreach (var arrayFormulaCell in ws.Range("A1:B2").Cells())
                {
                    Assert.AreEqual("1+2", arrayFormulaCell.FormulaA1);
                    Assert.AreEqual("A1:B2", arrayFormulaCell.FormulaReference.ToStringRelative());
                }

                var outsideCell = ws.Cell("A3");
                Assert.IsEmpty(outsideCell.FormulaA1);
                Assert.Null(outsideCell.FormulaReference);
            }, @"Other\Formulas\ArrayFormula.xlsx");
        }

        [Test]
        public void CanBeOnlyForOneCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var oneCell = ws.Cell("B3");

            oneCell.AsRange().FormulaArrayA1 = "2+5";

            Assert.True(oneCell.HasArrayFormula);
            Assert.AreEqual("2+5", oneCell.FormulaA1);
            Assert.AreEqual("B3:B3", oneCell.FormulaReference.ToStringRelative());
        }

        [Test]
        [Issue("112")]
        public void Array_formula_in_one_cell_in_the_first_cell_of_the_sheet_has_a_reference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("A1");

            cell.AsRange().FormulaArrayA1 = "2+5";

            Assert.AreEqual("A1:A1", cell.FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("112")]
        public void Array_formula_in_one_cell_in_the_first_cell_of_the_sheet_is_saved_and_loaded()
        {
            using var stream = new System.IO.MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet().Cell("A1").AsRange().FormulaArrayA1 = "2+5";
                wb.SaveAs(stream);
            }

            using var loaded = new XLWorkbook(stream);
            var cell = loaded.Worksheet(1).Cell("A1");
            Assert.True(cell.HasArrayFormula);
            Assert.AreEqual("2+5", cell.FormulaA1);
            Assert.AreEqual("A1:A1", cell.FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("112")]
        public void Array_formula_reference_cannot_be_set_to_null()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";
            var cell = ws.Cell("B2");

            Assert.Throws<ArgumentException>(() => cell.FormulaReference = null);
            Assert.AreEqual("B2:C3", cell.FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("112")]
        public void Data_table_reference_cannot_be_set_to_null()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);
            var cell = ws.Cell("B2");

            Assert.Throws<ArgumentException>(() => cell.FormulaReference = null);
            Assert.AreEqual("B2:C3", cell.FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("112")]
        public void Normal_formula_reference_can_be_set_to_null()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("B2");
            cell.FormulaA1 = "2+5";

            cell.FormulaReference = null;

            Assert.IsNull(cell.FormulaReference);
            Assert.AreEqual("2+5", cell.FormulaA1);
        }

        [TestCase("B2:C3")]
        [TestCase("B2:C4")]
        [TestCase("A1:D7")]
        public void SettingValueToContainingRangeClearsArrayFormula(string containingRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var arrayFormulaRange = ws.Range("B2:C3");
            arrayFormulaRange.FormulaArrayA1 = "5";

            ws.Range(containingRange).Value = Blank.Value;

            foreach (var cell in arrayFormulaRange.Cells())
            {
                Assert.AreEqual(Blank.Value, cell.Value);
                Assert.False(cell.HasArrayFormula);
                Assert.IsEmpty(cell.FormulaA1);
                Assert.Null(cell.FormulaReference);
            }
        }

        [TestCase("B2:D3")]
        [TestCase("A1:E4")]
        public void SettingFormulaToContainingRangeClearsOriginalArrayFormula(string overlapRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D3").FormulaArrayA1 = "1";

            Assert.DoesNotThrow(() => ws.Range(overlapRange).FormulaArrayA1 = "2");
        }

        [TestCase("B2:B2")]
        [TestCase("B2:B3")]
        [TestCase("A1:C3")]
        [TestCase("D2:F3")]
        [TestCase("C:C")]
        [TestCase("2:2")]
        public void ArrayFormulaCantPartiallyOverlapWithAnotherArrayFormula(string partialOverlapRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D3").FormulaArrayA1 = "1";

            Assert.That(() => ws.Range(partialOverlapRange).FormulaArrayA1 = "2",
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Can't create array function that partially covers another array function."));
        }

        [TestCase("A1:B2")]
        [TestCase("A2")]
        public void ArrayFormulaCantOverlapWithMergedRange(string partialOverlapRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:A2").Merge();

            Assert.That(() => ws.Range(partialOverlapRange).FormulaArrayA1 = "1",
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Can't create array function over a merged range."));
        }

        [TestCase("A1:B2")]
        [TestCase("A1:C1")]
        public void ArrayFormulaCantOverlapWithTable(string formulaRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("A2").Value = 5;
            ws.Range("A1:A2").CreateTable();

            Assert.That(() => ws.Range(formulaRange).FormulaArrayA1 = "1",
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Can't create array function over a table."));
        }

        [TestCase("A1:B3")]
        [TestCase("B1:C4")]
        [Issue("118")]
        public void Table_cannot_be_created_over_array_formula_with_more_than_one_cell(string tableRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            Assert.That(() => ws.Range(tableRange).CreateTable(),
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("The table range has cells of the array formula or data table in B2:C3. A table cannot contain an array formula or a data table with more than one cell."));

            Assert.IsEmpty(ws.Tables);
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("118")]
        public void Table_can_be_created_over_array_formula_in_one_cell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("A2").AsRange().FormulaArrayA1 = "2+5";

            var table = ws.Range("A1:A2").CreateTable();

            Assert.AreEqual("A1:A2", table.RangeAddress.ToStringRelative());
            Assert.True(ws.Cell("A2").HasArrayFormula);
        }

        [Test]
        [Issue("118")]
        public void Table_cannot_be_created_over_part_of_data_table()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

            // The data table has its formula only in its first cell, B2. The table range does not contain B2.
            Assert.Throws<InvalidOperationException>(() => ws.Range("C3:D4").CreateTable());
            Assert.IsEmpty(ws.Tables);
        }

        [Test]
        [Issue("118")]
        public void Table_resize_over_array_formula_with_more_than_one_cell_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Range("A2:A3").Value = 5;
            var table = ws.Range("A1:A3").CreateTable();
            ws.Range("C2:D3").FormulaArrayA1 = "2+5";

            Assert.Throws<InvalidOperationException>(() => table.Resize(ws.Range("A1:C3")));

            Assert.AreEqual("A1:A3", table.RangeAddress.ToStringRelative());
            Assert.AreEqual("C2:D3", ws.Cell("D3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("118")]
        public void Workbook_with_array_formula_in_table_is_loaded()
        {
            using var stream = new System.IO.MemoryStream();
            using (var wb = new XLWorkbook())
            {
                // Other producers can write a table with an array formula. The public methods do not allow it.
                var ws = (XLWorksheet)wb.AddWorksheet();
                ws.Cell("A1").Value = "Name";
                ws.Cell("B1").Value = "Value";
                ws.Range("A1:B3").CreateTable("Data");
                var arrayRange = Area.Parse("A2:B3");
                ws.Internals.CellsCollection.FormulaSlice.SetArray(arrayRange, XLCellFormula.Array("2+5", arrayRange, false));
                wb.SaveAs(stream);
            }

            using var loaded = new XLWorkbook(stream);
            var loadedSheet = loaded.Worksheet(1);
            Assert.AreEqual("A1:B3", loadedSheet.Table("Data").RangeAddress.ToStringRelative());
            Assert.AreEqual("A2:B3", loadedSheet.Cell("B3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("118")]
        public void Insert_of_table_over_array_formula_with_more_than_one_cell_throws_and_keeps_the_sheet()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";
            var data = new[] { new { Name = "a", Value = 1, Other = 2 }, new { Name = "b", Value = 3, Other = 4 } };

            Assert.That(() => ws.Cell("A1").InsertTable(data),
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("The table range has cells of the array formula or data table in B2:C3. A table cannot contain an array formula or a data table with more than one cell."));

            Assert.IsEmpty(ws.Tables);
            Assert.True(ws.Cell("A1").Value.IsBlank);
            Assert.True(ws.Cell("A2").Value.IsBlank);
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("118")]
        public void Insert_of_table_next_to_array_formula_creates_the_table()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("C2:D3").FormulaArrayA1 = "2+5";
            var data = new[] { new { Name = "a", Value = 1 }, new { Name = "b", Value = 3 } };

            var table = ws.Cell("A1").InsertTable(data);

            Assert.AreEqual("A1:B3", table.RangeAddress.ToStringRelative());
            Assert.AreEqual("b", ws.Cell("A3").GetString());
            Assert.AreEqual("C2:D3", ws.Cell("D3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("118")]
        public void Insert_of_data_without_table_over_array_formula_does_not_check_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";
            var data = new[] { new { Name = "a", Value = 1 } };

            var table = ws.Cell("A1").InsertTable(data, createTable: false);

            Assert.AreEqual("A1:B2", table.RangeAddress.ToStringRelative());
            Assert.IsEmpty(ws.Tables);
        }

        [Test]
        [Issue("118")]
        public void Table_copy_to_sheet_with_array_formula_throws()
        {
            using var wb = new XLWorkbook();
            var source = wb.AddWorksheet("Source");
            source.Cell("A1").Value = "Name";
            source.Cell("B1").Value = "Value";
            var table = source.Range("A1:B3").CreateTable("Data");
            var target = wb.AddWorksheet("Target");
            target.Range("B2:C3").FormulaArrayA1 = "2+5";

            Assert.Throws<InvalidOperationException>(() => table.CopyTo(target));

            Assert.IsEmpty(target.Tables);
            Assert.True(target.Cell("A1").Value.IsBlank);
            Assert.AreEqual("B2:C3", target.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("118")]
        public void Array_formula_in_one_cell_of_table_can_be_created()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("A2").Value = 5;
            ws.Range("A1:A2").CreateTable();

            ws.Cell("A2").AsRange().FormulaArrayA1 = "2+5";

            Assert.True(ws.Cell("A2").HasArrayFormula);
            Assert.AreEqual("A2:A2", ws.Cell("A2").FormulaReference?.ToStringRelative());
        }

        [Test]
        public void SettingArrayFormulaInvalidatesCells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.False(ws.Cell("A1").NeedsRecalculation);
            Assert.False(ws.Cell("A2").NeedsRecalculation);

            ws.Range("A1:A2").FormulaArrayA1 = "ABS(-3)";

            Assert.True(ws.Cell("A1").NeedsRecalculation);
            Assert.True(ws.Cell("A2").NeedsRecalculation);
        }

        [Test]
        public void ReferencingItselfIsCircularError()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "A2";
            ws.Range("A2").FormulaArrayA1 = "A1";

            Assert.That(() => _ = ws.Cell("A2").Value,
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("Formula in a cell Sheet1!A1 is part of a cycle."));
        }
    }
}
