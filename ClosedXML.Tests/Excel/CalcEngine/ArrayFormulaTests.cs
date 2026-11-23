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

        [TestCase("B2:B3", XLClearOptions.All)]
        [TestCase("B2:B3", XLClearOptions.Contents)]
        [TestCase("A1:B2", XLClearOptions.All)]
        [TestCase("C:C", XLClearOptions.Contents)]
        [Issue("115")]
        public void Clear_of_part_of_array_formula_throws_and_keeps_the_cells(string clearedRange, XLClearOptions clearOptions)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";
            ws.Cell("A1").Value = 4;

            Assert.That(() => ws.Range(clearedRange).Clear(clearOptions),
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("The operation changes only part of the array formula or data table in B2:C3. Change all cells of the range, or remove the formula first."));

            Assert.AreEqual(4, ws.Cell("A1").Value);
            foreach (var cell in ws.Range("B2:C3").Cells())
            {
                Assert.True(cell.HasArrayFormula);
                Assert.AreEqual("B2:C3", cell.FormulaReference?.ToStringRelative());
            }
        }

        [Test]
        [Issue("115")]
        public void Clear_of_one_cell_of_array_formula_throws_and_keeps_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            Assert.Throws<InvalidOperationException>(() => ws.Cell("C3").Clear());

            Assert.True(ws.Cell("C3").HasArrayFormula);
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [TestCase("B2:C3")]
        [TestCase("A1:D4")]
        [Issue("115")]
        public void Clear_of_all_cells_of_array_formula_removes_the_formula(string clearedRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            ws.Range(clearedRange).Clear(XLClearOptions.Contents);

            Assert.That(ws.Range("B2:C3").Cells(), Has.None.Matches<IXLCell>(cell => cell.HasFormula));
        }

        [Test]
        [Issue("115")]
        public void Clear_of_cell_with_array_formula_in_one_cell_removes_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B2").AsRange().FormulaArrayA1 = "2+5";

            ws.Cell("B2").Clear();

            Assert.False(ws.Cell("B2").HasFormula);
        }

        [Test]
        [Issue("115")]
        public void Clear_of_formats_of_part_of_array_formula_keeps_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";
            ws.Range("B2:C3").Style.Font.Bold = true;

            ws.Range("B2:B3").Clear(XLClearOptions.AllFormats);
            ws.Cell("C3").Clear(XLClearOptions.AllFormats);

            Assert.False(ws.Cell("B2").Style.Font.Bold);
            Assert.False(ws.Cell("C3").Style.Font.Bold);
            Assert.True(ws.Cell("C2").Style.Font.Bold);
            Assert.True(ws.Cell("B2").HasArrayFormula);
            Assert.AreEqual("B2:C3", ws.Cell("B2").FormulaReference?.ToStringRelative());
        }

        private static readonly object[] PartialCollectionClears =
        {
            new object[] { "cells", (Action<IXLWorksheet>)(ws => ws.Cells("A1,C3").Clear()) },
            new object[] { "ranges", (Action<IXLWorksheet>)(ws => ws.Ranges("A1,B2:B3").Clear()) },
            new object[] { "rows", (Action<IXLWorksheet>)(ws => ws.Rows("1,3").Clear()) },
            new object[] { "columns", (Action<IXLWorksheet>)(ws => ws.Columns("A,C").Clear()) },
            new object[] { "range rows", (Action<IXLWorksheet>)(ws => ws.Range("A1:D4").Rows(1, 2).Clear()) },
            new object[] { "range columns", (Action<IXLWorksheet>)(ws => ws.Range("A1:D4").Columns(1, 2).Clear()) },
            new object[] { "table rows", (Action<IXLWorksheet>)(ws => ws.Table("Data").DataRange.Rows(1, 1).Clear()) },
            new object[] { "tables", (Action<IXLWorksheet>)(ws => ws.Tables.Clear()) },
        };

        [TestCaseSource(nameof(PartialCollectionClears))]
        [Issue("115")]
        public void Clear_of_collection_checks_all_parts_before_the_first_clear(string collection, Action<IXLWorksheet> clear)
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            ws.Cell("A1").Value = "Header";
            ws.Cell("B1").Value = "Result";
            ws.Range("A2:A3").Value = 4;
            ws.Range("A1:B3").CreateTable("Data");

            // The public methods do not allow a table over an array formula. A file from another producer can have it.
            var arrayRange = Area.Parse("B2:C3");
            ws.Internals.CellsCollection.FormulaSlice.SetArray(arrayRange, XLCellFormula.Array("2+5", arrayRange, false));

            Assert.Throws<InvalidOperationException>(() => clear(ws), collection);

            Assert.AreEqual("Header", ws.Cell("A1").Value, collection);
            Assert.AreEqual(4, ws.Cell("A2").Value, collection);
            Assert.True(ws.Cell("C3").HasArrayFormula, collection);
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative(), collection);
        }

        private static readonly object[] CollectionClearsOutsideArrayFormula =
        {
            new object[] { "range rows", (Action<IXLWorksheet>)(ws => ws.Range("A2:A3").Rows(1, 2).Clear()) },
            new object[] { "range columns", (Action<IXLWorksheet>)(ws => ws.Range("A2:A3").Columns(1, 1).Clear()) },
            new object[] { "table rows", (Action<IXLWorksheet>)(ws => ws.Table("Data").DataRange.Rows(1, 2).Clear()) },
            new object[] { "tables", (Action<IXLWorksheet>)(ws => ws.Tables.Clear()) },
        };

        [TestCaseSource(nameof(CollectionClearsOutsideArrayFormula))]
        [Issue("115")]
        public void Clear_of_collection_outside_array_formula_clears_the_cells(string collection, Action<IXLWorksheet> clear)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Header";
            ws.Range("A2:A3").Value = 4;
            ws.Range("A1:A3").CreateTable("Data");
            ws.Range("C2:D3").FormulaArrayA1 = "2+5";

            clear(ws);

            Assert.AreEqual(Blank.Value, ws.Cell("A2").Value, collection);
            Assert.AreEqual(Blank.Value, ws.Cell("A3").Value, collection);
            Assert.AreEqual("C2:D3", ws.Cell("D3").FormulaReference?.ToStringRelative(), collection);
        }

        [Test]
        [Issue("115")]
        public void Clear_of_merged_cell_in_collection_checks_the_merged_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B1").FormulaArrayA1 = "2+5";

            // Merge does not allow an array formula in merged cells. A file from another producer can have it.
            ws.Range("A1:A2").Merge(false);

            // The clear of A2 clears the merged range A1:A2. A1 is a cell of the array formula.
            Assert.Throws<InvalidOperationException>(() => ws.Cells("A2").Clear());

            Assert.AreEqual("A1:B1", ws.Cell("A1").FormulaReference?.ToStringRelative());
            Assert.AreEqual("A1:B1", ws.Cell("B1").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("115")]
        public void Clear_of_collection_with_all_cells_of_array_formula_removes_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            ws.Cells("B2:C3").Clear(XLClearOptions.Contents);

            Assert.That(ws.Range("B2:C3").Cells(), Has.None.Matches<IXLCell>(cell => cell.HasFormula));
        }

        [Test]
        [Issue("115")]
        public void Clear_of_ranges_that_together_contain_the_array_formula_removes_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            ws.Ranges("B2:C2,B2:B3,C3").Clear(XLClearOptions.Contents);

            Assert.That(ws.Range("B2:C3").Cells(), Has.None.Matches<IXLCell>(cell => cell.HasFormula));
        }

        [Test]
        [Issue("115")]
        public void Clear_of_part_of_data_table_throws()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

            // The data table has its formula only in the first cell, so the clear starts in another cell.
            Assert.Throws<InvalidOperationException>(() => ws.Range("C3:D4").Clear());
            Assert.Throws<InvalidOperationException>(() => ws.Cell("C3").Clear());

            Assert.AreSame(formula, ws.Internals.CellsCollection.FormulaSlice.Get(new Point(2, 2)));
        }

        [Test]
        [Issue("115")]
        public void Clear_of_part_of_data_table_after_insert_throws()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

            ws.Row(1).InsertRowsAbove(2);

            Assert.AreEqual("B4:C5", formula.Range.ToString());
            Assert.Throws<InvalidOperationException>(() => ws.Cell("C5").Clear());
            Assert.DoesNotThrow(() => ws.Cell("C3").Clear());
        }

        [Test]
        [Issue("115")]
        public void Clear_after_removal_of_data_table_does_not_throw()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);
            ws.Cell("C3").Value = 5;

            ws.Range("B2:C3").Clear();
            ws.Cell("C3").Value = 6;

            Assert.DoesNotThrow(() => ws.Cell("C3").Clear());
            Assert.IsEmpty(ws.Internals.CellsCollection.FormulaSlice.GetDataTables());
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

        [TestCase("B2:C2")]
        [TestCase("B2:C3")]
        [TestCase("A1:D4")]
        [Issue("119")]
        public void Merge_over_array_formula_throws_and_keeps_the_sheet(string mergeRange)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            Assert.That(() => ws.Range(mergeRange).Merge(),
                Throws.TypeOf<InvalidOperationException>()
                    .With.Message.EqualTo("The merge range has cells of the array formula or data table in B2:C3. Merged cells cannot contain an array formula or a data table."));

            Assert.IsEmpty(ws.MergedRanges);
            Assert.AreEqual("B2:C3", ws.Cell("C2").FormulaReference?.ToStringRelative());
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("119")]
        public void Merge_over_array_formula_in_one_cell_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").AsRange().FormulaArrayA1 = "2+5";

            Assert.Throws<InvalidOperationException>(() => ws.Range("A1:B1").Merge());

            Assert.IsEmpty(ws.MergedRanges);
            Assert.True(ws.Cell("A1").HasArrayFormula);
        }

        [Test]
        [Issue("119")]
        public void Merge_over_part_of_data_table_throws()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

            // The data table has its formula only in its first cell, B2. The merge range does not contain B2.
            Assert.Throws<InvalidOperationException>(() => ws.Range("C3:D3").Merge());
            Assert.IsEmpty(ws.MergedRanges);
        }

        [Test]
        [Issue("119")]
        public void Merge_next_to_array_formula_merges_the_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";

            ws.Range("D2:E3").Merge();

            Assert.AreEqual("D2:E3", ws.MergedRanges.Single().RangeAddress.ToStringRelative());
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("119")]
        public void Workbook_with_array_formula_in_merged_range_is_loaded()
        {
            using var stream = new System.IO.MemoryStream();
            using (var wb = new XLWorkbook())
            {
                // Other producers can write a merged range over an array formula. Merge(false) does not check it.
                var ws = wb.AddWorksheet();
                ws.Range("B2:C3").FormulaArrayA1 = "2+5";
                ws.Range("B2:C2").Merge(false);
                wb.SaveAs(stream);
            }

            using var loaded = new XLWorkbook(stream);
            var loadedSheet = loaded.Worksheet(1);
            Assert.AreEqual("B2:C2", loadedSheet.MergedRanges.Single().RangeAddress.ToStringRelative());
            Assert.AreEqual("B2:C3", loadedSheet.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("125")]
        public void Sheet_copy_keeps_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 3;
            ws.Range("B2:C3").FormulaArrayA1 = "A1*2";

            var copy = ws.CopyTo("Copy");

            Assert.True(copy.Cell("C3").HasArrayFormula);
            Assert.AreEqual("A1*2", copy.Cell("C3").FormulaA1);
            Assert.AreEqual("B2:C3", copy.Cell("C3").FormulaReference?.ToStringRelative());

            // The cells of the copy share one formula, and the source keeps its own formula.
            var copyFormula = ((XLCell)copy.Cell("B2")).Formula;
            Assert.AreSame(copyFormula, ((XLCell)copy.Cell("C3")).Formula);
            Assert.AreNotSame(((XLCell)ws.Cell("B2")).Formula, copyFormula);

            copy.Cell("A1").Value = 5;
            Assert.AreEqual(10, copy.Cell("C3").Value);
            Assert.AreEqual(6, ws.Cell("C3").Value);
        }

        [Test]
        [Issue("125")]
        public void Sheet_copy_keeps_data_table()
        {
            using var stream = new System.IO.MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = (XLWorksheet)wb.AddWorksheet("Source");
                var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
                ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

                var copy = (XLWorksheet)ws.CopyTo("Copy");

                var copyFormula = copy.Internals.CellsCollection.FormulaSlice.Get(new Point(2, 2));
                Assert.NotNull(copyFormula);
                Assert.AreNotSame(formula, copyFormula);
                Assert.AreEqual(FormulaType.DataTable, copyFormula!.Type);
                Assert.AreEqual(Area.Parse("B2:C3"), copyFormula.Range);
                Assert.AreSame(copyFormula, copy.Internals.CellsCollection.FormulaSlice.GetDataTables().Single());
                wb.SaveAs(stream);
            }

            using var loaded = new XLWorkbook(stream);
            var loadedCopy = (XLWorksheet)loaded.Worksheet("Copy");
            var loadedFormula = loadedCopy.Internals.CellsCollection.FormulaSlice.GetDataTables().Single();
            Assert.AreEqual(Area.Parse("B2:C3"), loadedFormula.Range);
            Assert.True(loadedFormula.IsRowDataTable);
        }

        [Test]
        [Issue("125")]
        public void Sheet_copy_keeps_array_formula_in_merged_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Other producers can write a merged range over an array formula.
            ws.Range("B2:C3").FormulaArrayA1 = "2+5";
            ws.Range("B2:C2").Merge(false);

            var copy = ws.CopyTo("Copy");

            Assert.AreEqual("B2:C2", copy.MergedRanges.Single().RangeAddress.ToStringRelative());
            Assert.AreEqual("B2:C3", copy.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("125")]
        public void Range_copy_keeps_array_formula_and_moves_relative_references()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "$A$1+A1";

            ws.Range("B2:C3").CopyTo(ws.Cell("E5"));

            foreach (var cell in ws.Range("E5:F6").Cells())
            {
                Assert.True(cell.HasArrayFormula);
                Assert.AreEqual("$A$1+D4", cell.FormulaA1);
                Assert.AreEqual("E5:F6", cell.FormulaReference?.ToStringRelative());
            }

            Assert.AreEqual("$A$1+A1", ws.Cell("C3").FormulaA1);
            Assert.AreEqual("B2:C3", ws.Cell("C3").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("125")]
        public void Range_copy_to_other_workbook_keeps_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "A1+1";
            using var targetWb = new XLWorkbook();
            var target = targetWb.AddWorksheet();

            ws.Range("A1:C3").CopyTo(target.Cell("A1"));

            Assert.AreEqual("A1+1", target.Cell("C3").FormulaA1);
            Assert.AreEqual("B2:C3", target.Cell("C3").FormulaReference?.ToStringRelative());
            target.Cell("A1").Value = 4;
            Assert.AreEqual(5, target.Cell("C3").Value);
        }

        [Test]
        [Issue("125")]
        public void Range_copy_of_part_of_array_formula_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "1+1";
            ws.Cell("E5").Value = "Target";

            Assert.That(() => ws.Range("B2:B3").CopyTo(ws.Cell("E5")),
                Throws.InvalidOperationException.With.Message.EqualTo("The copy has only part of the array formula in B2:C3. Copy all cells of the range."));

            Assert.AreEqual("Target", ws.Cell("E5").Value);
            Assert.False(ws.Cell("E6").HasFormula);
        }

        [Test]
        [Issue("125")]
        public void Range_copy_of_data_table_copies_only_the_values()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            ws.Cell("B2").Value = 7;
            ws.Cell("C3").Value = 8;
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

            ws.Range("B2:C3").CopyTo(ws.Cell("E5"));

            Assert.False(ws.Cell("E5").HasFormula);
            Assert.AreEqual(7, ws.Cell("E5").Value);
            Assert.AreEqual(8, ws.Cell("F6").Value);
            Assert.AreSame(formula, ws.Internals.CellsCollection.FormulaSlice.GetDataTables().Single());
        }

        [Test]
        [Issue("125")]
        public void Cell_copy_of_array_formula_in_one_cell_keeps_it()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2").FormulaArrayA1 = "A1+1";

            ws.Cell("D4").CopyFrom(ws.Cell("B2"));

            Assert.True(ws.Cell("D4").HasArrayFormula);
            Assert.AreEqual("C3+1", ws.Cell("D4").FormulaA1);
            Assert.AreEqual("D4:D4", ws.Cell("D4").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("125")]
        public void Cell_copy_of_part_of_array_formula_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "1+1";

            Assert.Throws<InvalidOperationException>(() => ws.Cell("E5").CopyFrom(ws.Cell("C3")));
            Assert.Throws<InvalidOperationException>(() => ws.Cell("C3").CopyTo(ws.Cell("E5")));

            Assert.False(ws.Cell("E5").HasFormula);
        }

        [Test]
        [Issue("125")]
        public void Cell_copy_to_part_of_array_formula_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "1+1";
            ws.Cell("A1").Value = 5;

            Assert.Throws<InvalidOperationException>(() => ws.Cell("C3").CopyFrom(ws.Cell("A1")));
            Assert.True(ws.Cell("C3").HasArrayFormula);

            // A copy of only the format does not change the contents of the cell.
            ((XLCell)ws.Cell("C3")).CopyFrom(ws.Cell("A1"), XLCellCopyOptions.Styles);
            Assert.True(ws.Cell("C3").HasArrayFormula);
        }

        [Test]
        [Issue("125")]
        public void Row_copy_keeps_array_formula_in_the_row()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C2").FormulaArrayA1 = "A2*3";

            ws.Row(2).CopyTo(ws.Row(10));

            Assert.AreEqual("A10*3", ws.Cell("C10").FormulaA1);
            Assert.AreEqual("B10:C10", ws.Cell("C10").FormulaReference?.ToStringRelative());
        }

        [Test]
        [Issue("125")]
        public void Row_copy_of_part_of_array_formula_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "1+1";
            ws.Cell("A10").Value = "Target";

            Assert.Throws<InvalidOperationException>(() => ws.Row(2).CopyTo(ws.Row(10)));

            Assert.AreEqual("Target", ws.Cell("A10").Value);
        }

        [Test]
        [Issue("125")]
        public void Column_copy_of_part_of_array_formula_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "1+1";
            ws.Cell("J1").Value = "Target";

            Assert.Throws<InvalidOperationException>(() => ws.Column(2).CopyTo(ws.Column(10)));

            Assert.AreEqual("Target", ws.Cell("J1").Value);
        }

        [Test]
        [Issue("125")]
        public void Range_copy_of_array_formula_over_table_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Value";
            ws.Cell("A2").Value = "Target";
            ws.Range("A1:B3").CreateTable();
            ws.Range("E5:F6").FormulaArrayA1 = "1+1";

            var ex = Assert.Throws<InvalidOperationException>(() => ws.Range("E5:F6").CopyTo(ws.Cell("A2")));

            StringAssert.Contains("A2:B3", ex!.Message);
            Assert.AreEqual("Target", ws.Cell("A2").Value);
            Assert.False(ws.Cell("A2").HasArrayFormula);
        }

        [Test]
        [Issue("125")]
        public void Range_copy_of_array_formula_in_one_cell_to_table_keeps_it()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Value";
            ws.Range("A1:B3").CreateTable();
            ws.Range("E5").FormulaArrayA1 = "1+1";

            ws.Range("E5").CopyTo(ws.Cell("A2"));

            Assert.True(ws.Cell("A2").HasArrayFormula);
            Assert.AreEqual("A2:A2", ws.Cell("A2").FormulaReference.ToString());
        }

        [Test]
        [Issue("125")]
        public void Row_copy_of_array_formula_over_table_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A4").Value = "Name";
            ws.Cell("B4").Value = "Value";
            ws.Cell("A5").Value = "Target";
            ws.Range("A4:B6").CreateTable();
            ws.Range("A1:B1").FormulaArrayA1 = "1+1";

            Assert.Throws<InvalidOperationException>(() => ws.Row(1).CopyTo(ws.Row(5)));

            Assert.AreEqual("Target", ws.Cell("A5").Value);
        }

        [Test]
        [Issue("125")]
        public void Cell_copy_of_array_formula_to_merged_cell_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1").FormulaArrayA1 = "1+1";
            ws.Range("C3:D3").Merge();
            ws.Cell("C3").Value = "Target";

            var ex = Assert.Throws<InvalidOperationException>(() => ws.Cell("C3").CopyFrom(ws.Cell("A1")));

            StringAssert.Contains("C3", ex!.Message);
            Assert.AreEqual("Target", ws.Cell("C3").Value);
        }

        [Test]
        [Issue("125")]
        public void Range_copy_of_array_formula_after_last_column_throws_and_keeps_the_target()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B1").FormulaArrayA1 = "1+1";
            ws.Cell("XFD3").Value = "Target";

            Assert.Throws<InvalidOperationException>(() => ws.Range("A1:B1").CopyTo(ws.Cell("XFD3")));

            Assert.AreEqual("Target", ws.Cell("XFD3").Value);
            Assert.False(ws.Cell("XFD3").HasArrayFormula);
        }

        [Test]
        [Issue("125")]
        public void Sheet_copy_keeps_table_with_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet("Source");
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Value";
            ws.Range("A1:B3").CreateTable();

            // A loaded file can have an array formula in a table. The public setter does not allow it.
            ws.Internals.CellsCollection.FormulaSlice.SetArray(Area.Parse("A2:B3"), XLCellFormula.Array("1+1", Area.Parse("A2:B3"), false));

            var copy = ws.CopyTo("Copy");

            Assert.AreEqual(1, copy.Tables.Count());
            Assert.AreEqual("A1:B3", copy.Tables.Single().RangeAddress.ToString());
            Assert.AreEqual("Name", copy.Cell("A1").Value);
            Assert.True(copy.Cell("B3").HasArrayFormula);
            Assert.AreEqual("A2:B3", copy.Cell("B3").FormulaReference.ToString());
        }

        [Test]
        [Issue("125")]
        public void Table_copy_keeps_array_formula_in_the_table()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet("Source");
            ws.Cell("A1").Value = "Name";
            ws.Cell("B1").Value = "Value";
            var table = ws.Range("A1:B3").CreateTable();

            // A loaded file can have an array formula in a table. The public setter does not allow it.
            ws.Internals.CellsCollection.FormulaSlice.SetArray(Area.Parse("A2:B3"), XLCellFormula.Array("1+1", Area.Parse("A2:B3"), false));
            var target = wb.AddWorksheet("Target");

            table.CopyTo(target);

            Assert.AreEqual("A1:B3", target.Tables.Single().RangeAddress.ToString());
            Assert.True(target.Cell("B3").HasArrayFormula);
            Assert.AreEqual("A2:B3", target.Cell("B3").FormulaReference!.ToString());
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
