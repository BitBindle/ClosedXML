using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Linq;

namespace ClosedXML.Tests.Excel
{
    [TestFixture]
    public class FormulaTests
    {
        [Test]
        public void CopyFormula()
        {
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").FormulaA1 = "B1";
            ws.Cell("A1").CopyTo("A2");
            Assert.AreEqual("B2", ws.Cell("A2").FormulaA1);
        }

        [Test]
        public void CopyFormula2()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").FormulaA1 = "A2-1";
                ws.Cell("A1").CopyTo("B1");
                Assert.AreEqual("R[1]C-1", ws.Cell("A1").FormulaR1C1);
                Assert.AreEqual("R[1]C-1", ws.Cell("B1").FormulaR1C1);
                Assert.AreEqual("B2-1", ws.Cell("B1").FormulaA1);

                ws.Cell("A1").FormulaA1 = "B1+1";
                ws.Cell("A1").CopyTo("A2");
                Assert.AreEqual("RC[1]+1", ws.Cell("A1").FormulaR1C1);
                Assert.AreEqual("RC[1]+1", ws.Cell("A2").FormulaR1C1);
                Assert.AreEqual("B2+1", ws.Cell("A2").FormulaA1);
            }
        }

        [Test]
        public void CopyFormulaWithSheetNameThatResemblesFormula()
        {
            using (var wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("S10 Data");
                ws.Cell("A1").Value = "Some value";
                ws.Cell("A2").Value = 123;

                ws = wb.Worksheets.Add("Summary");
                ws.Cell("A1").FormulaA1 = "='S10 Data'!A1";
                Assert.AreEqual("Some value", ws.Cell("A1").Value);

                ws.Cell("A1").CopyTo("A2");
                Assert.AreEqual("'S10 Data'!A2", ws.Cell("A2").FormulaA1);

                ws.Cell("A1").CopyTo("B1");
                Assert.AreEqual("'S10 Data'!B1", ws.Cell("B1").FormulaA1);

                ws.Cell("A3").FormulaA1 = "=SUM('S10 Data'!A2)";
                Assert.AreEqual(123, ws.Cell("A3").Value);
            }
        }

        [Test]
        public void FormulaWithReferenceIncludingSheetName()
        {
            using (var wb = new XLWorkbook())
            {
                object value;
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
                ws.Cell("B1").FormulaA1 = "=SUM(A1:A50)";
                value = ws.Cell("B1").Value;
                Assert.AreEqual(1275, value);

                ws = wb.AddWorksheet("Sheet2");

                ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet1!A50)";
                value = ws.Cell("A1").Value;
                Assert.AreEqual(1275, value);

                ws.Cell("B1").FormulaA1 = "=SUM(Sheet1!A1:A50)";
                value = ws.Cell("B1").Value;
                Assert.AreEqual(1275, value);
            }
        }

        [Test]
        public void InvalidReferences()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").InsertData(Enumerable.Range(1, 50));
                ws = wb.AddWorksheet("Sheet2");

                ws.Cell("A1").FormulaA1 = "=SUM(Sheet1!A1:Sheet2!A50)";
                Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("A1").Value);

                ws.Cell("B1").FormulaA1 = "=SUM(UnknownSheet!A50)";
                Assert.AreEqual(XLError.CellReference, ws.Cell("B1").Value);
            }
        }

        [Test]
        public void DateAgainstStringComparison()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("A1").Value = new DateTime(2016, 1, 1);

                ws.Cell("A2").FormulaA1 = @"=IF(A1 = """", ""A"", ""B"")";
                var actual = ws.Cell("A2").Value;
                Assert.AreEqual(actual, "B");

                ws.Cell("A3").FormulaA1 = @"=IF("""" = A1, ""A"", ""B"")";
                actual = ws.Cell("A3").Value;
                Assert.AreEqual(actual, "B");
            }
        }

        [Test]
        public void FormulaThatReferencesEntireRow()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().Value = 1;
                ws.FirstCell().CellRight().Value = 2;
                ws.FirstCell().CellRight(5).Value = 3;

                ws.FirstCell().CellBelow().FormulaA1 = "=SUM(1:1)";

                var actual = ws.FirstCell().CellBelow().Value;
                Assert.AreEqual(6, actual);
            }
        }

        [Test]
        public void FormulaThatReferencesEntireColumn()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().Value = 1;
                ws.FirstCell().CellBelow().Value = 2;
                ws.FirstCell().CellBelow(5).Value = 3;

                ws.FirstCell().CellRight().FormulaA1 = "=SUM(A:A)";

                var actual = ws.FirstCell().CellRight().Value;
                Assert.AreEqual(6, actual);
            }
        }

        [Test]
        public void FormulaThatStartsWithEqualsAndPlus()
        {
            object actual;
            actual = XLWorkbook.EvaluateExpr("=MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);

            actual = XLWorkbook.EvaluateExpr("=+MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);

            actual = XLWorkbook.EvaluateExpr("=+++++MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);

            actual = XLWorkbook.EvaluateExpr("+MID(\"This is a test\", 6, 2)");
            Assert.AreEqual("is", actual);
        }

        [Test]
        public void UnimplementedStandardFunctionsAreEvaluatedToNameNotFoundError()
        {
            // RTD will never be implemented
            var actual = XLWorkbook.EvaluateExpr("RTD(\"MyRTDServerProdID\",\"MyServer\",\"RaceNum\",\"RunnerID\",\"StatType\")");
            Assert.AreEqual(XLError.NameNotRecognized, actual);
        }

        [Test]
        public void FormulasWithErrors()
        {
            Assert.AreEqual(XLError.CellReference, XLWorkbook.EvaluateExpr("YEAR(#REF!)"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("YEAR(#VALUE!)"));
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("YEAR(#DIV/0!)"));
            Assert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("YEAR(#NAME?)"));
            Assert.AreEqual(XLError.NoValueAvailable, XLWorkbook.EvaluateExpr("YEAR(#N/A)"));
            Assert.AreEqual(XLError.NullValue, XLWorkbook.EvaluateExpr("YEAR(#NULL!)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("YEAR(#NUM!)"));
        }

        [Test]
        public void LegacyFunctionPropagateErrorWithoutException()
        {
            Assert.AreEqual(XLError.NameNotRecognized, XLWorkbook.EvaluateExpr("SIN(YEAR(#NAME?))+1"));
        }

        [Test]
        public void UnicodeLetterParsing()
        {
            using (var wb = new XLWorkbook())
            {
                var ws1 = wb.AddWorksheet("Sheet C CÄ");
                var ws2 = wb.AddWorksheet("ÖC");
                var ws3 = wb.AddWorksheet("Sheet3");

                ws1.FirstCell().SetValue(100);
                ws2.FirstCell().SetValue(50);

                ws3.FirstCell().FormulaA1 = "='Sheet C CÄ'!A1";
                ws3.FirstCell().CellBelow().FormulaA1 = "ÖC!A1";

                Assert.AreEqual(100, ws3.FirstCell().Value);
                Assert.AreEqual(50, ws3.FirstCell().CellBelow().Value);
            }
        }

        [Test]
        [Issue("88")]
        public void ShiftFormula()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.Cell("B1").FormulaA1 = "ATAN2(C1,C2)";
                ws.Cell("B2").FormulaA1 = "DEC2HEX(C2)";
                ws.Range("B3:B5").FormulaArrayA1 = "DAYS360(C3:C5, D3:D5)";

                ws.Column(1).Delete();

                Assert.AreEqual("ATAN2(B1,B2)", ws.Cell("A1").FormulaA1);
                Assert.AreEqual("DEC2HEX(B2)", ws.Cell("A2").FormulaA1);
                Assert.True(ws.Cell("A3").HasArrayFormula);
                Assert.AreEqual("DAYS360(B3:B5, C3:C5)", ws.Cell("A3").FormulaA1);
            }
        }

        [Test]
        [Issue("88")]
        public void Array_formula_stays_an_array_formula_after_an_insert_below_it()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A1:A2").FormulaArrayA1 = "B1:B2*C1:C2";

            ws.Row(5).InsertRowsAbove(1);

            Assert.True(ws.Cell("A1").HasArrayFormula);
            Assert.True(ws.Cell("A2").HasArrayFormula);
            Assert.AreEqual("B1:B2*C1:C2", ws.Cell("A1").FormulaA1);
        }

        [Test]
        [Issue("88")]
        public void Insert_of_a_row_above_an_array_formula_moves_the_array_formula()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                ws.Cell("B1").Value = 2;
                ws.Cell("B2").Value = 3;
                ws.Cell("C1").Value = 5;
                ws.Cell("C2").Value = 7;
                ws.Range("A1:A2").FormulaArrayA1 = "B1:B2*C1:C2";

                ws.Row(1).InsertRowsAbove(1);

                Assert.False(ws.Cell("A1").HasFormula);
                Assert.True(ws.Cell("A2").HasArrayFormula);
                Assert.True(ws.Cell("A3").HasArrayFormula);
                Assert.AreEqual("B2:B3*C2:C3", ws.Cell("A2").FormulaA1);
                Assert.AreEqual("A2:A3", ws.Cell("A3").FormulaReference!.ToString());
                Assert.AreEqual(10, ws.Cell("A2").Value);
                Assert.AreEqual(21, ws.Cell("A3").Value);
                wb.SaveAs(ms);
            }

            using var loaded = new XLWorkbook(ms);
            var loadedSheet = loaded.Worksheet("Sheet1");
            Assert.True(loadedSheet.Cell("A3").HasArrayFormula);
            Assert.AreEqual("B2:B3*C2:C3", loadedSheet.Cell("A3").FormulaA1);
            Assert.AreEqual("A2:A3", loadedSheet.Cell("A2").FormulaReference!.ToString());
        }

        [Test]
        [Issue("88")]
        public void Insert_in_another_sheet_changes_only_the_references_of_an_array_formula()
        {
            using var wb = new XLWorkbook();
            var data = wb.AddWorksheet("Data");
            var report = wb.AddWorksheet("Report");
            report.Range("A1:A2").FormulaArrayA1 = "Data!B1:B2*2";

            data.Row(1).InsertRowsAbove(1);

            Assert.True(report.Cell("A1").HasArrayFormula);
            Assert.AreEqual("Data!B2:B3*2", report.Cell("A2").FormulaA1);
            Assert.AreEqual("A1:A2", report.Cell("A1").FormulaReference!.ToString());
        }

        [Test]
        [Issue("88")]
        public void Unparsable_array_formula_stays_the_same_after_an_insert()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Range("A2:A3").FormulaArrayA1 = "B2:B3*";

            ws.Row(1).InsertRowsAbove(1);

            Assert.True(ws.Cell("A3").HasArrayFormula);
            Assert.AreEqual("B2:B3*", ws.Cell("A3").FormulaA1);
        }

        [Test]
        [Issue("88")]
        public void Insert_of_a_column_moves_the_range_and_the_inputs_of_a_data_table()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
            var formula = XLCellFormula.DataTable2D(Area.Parse("B2:C3"), new Point(1, 1), false, new Point(2, 1), false);
            ws.Internals.CellsCollection.FormulaSlice.SetArray(formula.Range, formula);

            ws.Column(1).InsertColumnsBefore(1);

            var shifted = ws.Cell("C2").Formula!;
            Assert.AreSame(shifted, ws.Cell("D3").Formula);
            Assert.AreEqual(FormulaType.DataTable, shifted.Type);
            Assert.AreEqual("C2:D3", shifted.Range.ToString());
            Assert.AreEqual(new Point(1, 2), shifted.Input1);
            Assert.AreEqual(new Point(2, 2), shifted.Input2);
            Assert.AreEqual("{TABLE(B2,B1}", shifted.A1);
        }

        [TestCase(true, "{TABLE(#REF!,}")]
        [TestCase(false, "{TABLE(,#REF!}")]
        [Issue("88")]
        public void Delete_of_the_input_of_a_data_table_marks_the_input_as_deleted(bool isRowDataTable, string expectedText)
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C3"), new Point(1, 1), false, isRowDataTable);
            ws.Internals.CellsCollection.FormulaSlice.SetArray(formula.Range, formula);

            ws.Row(1).Delete();

            var shifted = ws.Cell("B1").Formula!;
            Assert.AreEqual(FormulaType.DataTable, shifted.Type);
            Assert.AreEqual("B1:C2", shifted.Range.ToString());
            Assert.True(shifted.Input1Deleted);
            Assert.AreEqual(expectedText, shifted.A1);
        }

        [Test]
        [Issue("88")]
        public void Delete_of_a_row_marks_the_second_input_of_a_data_table_as_deleted()
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet("Sheet1");
            var formula = XLCellFormula.DataTable2D(Area.Parse("B3:C4"), new Point(1, 1), false, new Point(2, 1), false);
            ws.Internals.CellsCollection.FormulaSlice.SetArray(formula.Range, formula);

            ws.Row(2).Delete();

            var shifted = ws.Cell("B2").Formula!;
            Assert.False(shifted.Input1Deleted);
            Assert.True(shifted.Input2Deleted);
            Assert.AreEqual("{TABLE(#REF!,A1}", shifted.A1);
        }

        [TestCase("B1", XLShiftDeletedCells.ShiftCellsUp)]
        [TestCase("A2", XLShiftDeletedCells.ShiftCellsLeft)]
        [TestCase("A1:A2", XLShiftDeletedCells.ShiftCellsUp)]
        [TestCase("A2:A3", XLShiftDeletedCells.ShiftCellsLeft)]
        [TestCase("A3:B3", XLShiftDeletedCells.ShiftCellsUp)]
        [Issue("102")]
        public void Delete_of_cells_that_changes_part_of_an_array_formula_throws(string deletedRange, XLShiftDeletedCells shift)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A2:B3").FormulaArrayA1 = "C2:D3*2";
            ws.Cell("A1").Value = 1;
            ws.Cell("B1").Value = 2;

            var ex = Assert.Throws<InvalidOperationException>(() => ws.Range(deletedRange).Delete(shift))!;

            StringAssert.Contains("A2:B3", ex.Message);
            AssertArrayFormula(ws, "A2:B3", "C2:D3*2");
            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
        }

        [TestCase("A2", true)]
        [TestCase("B1", false)]
        [Issue("102")]
        public void Insert_of_cells_that_moves_part_of_an_array_formula_throws(string cell, bool shiftDown)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B2").FormulaArrayA1 = "C1:D2*2";

            Assert.Throws<InvalidOperationException>(() =>
            {
                if (shiftDown)
                    ws.Cell(cell).InsertCellsAbove(1);
                else
                    ws.Cell(cell).InsertCellsBefore(1);
            });

            AssertArrayFormula(ws, "A1:B2", "C1:D2*2");
        }

        [Test]
        [Issue("102")]
        public void Insert_of_a_row_inside_an_array_formula_throws_and_keeps_the_rows()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:A3").FormulaArrayA1 = "B1:B3*2";
            ws.Row(3).Height = 30;

            Assert.Throws<InvalidOperationException>(() => ws.Row(2).InsertRowsAbove(1));
            Assert.Throws<InvalidOperationException>(() => ws.Row(1).InsertRowsBelow(1));

            AssertArrayFormula(ws, "A1:A3", "B1:B3*2");
            Assert.AreEqual(30, ws.Row(3).Height);
            Assert.AreEqual(ws.RowHeight, ws.Row(4).Height);
        }

        [Test]
        [Issue("102")]
        public void Insert_of_a_column_inside_an_array_formula_throws_and_keeps_the_columns()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:C1").FormulaArrayA1 = "A2:C2*2";
            ws.Column(3).Width = 30;

            Assert.Throws<InvalidOperationException>(() => ws.Column(2).InsertColumnsBefore(1));
            Assert.Throws<InvalidOperationException>(() => ws.Column(1).InsertColumnsAfter(1));

            AssertArrayFormula(ws, "A1:C1", "A2:C2*2");
            Assert.AreEqual(30, ws.Column(3).Width);
            Assert.AreEqual(ws.ColumnWidth, ws.Column(4).Width);
        }

        [Test]
        [Issue("102")]
        public void Insert_of_a_column_before_an_array_formula_moves_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:C1").FormulaArrayA1 = "A2:C2*2";

            ws.Column(1).InsertColumnsBefore(1);

            AssertArrayFormula(ws, "B1:D1", "B2:D2*2");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_a_row_or_a_column_inside_an_array_formula_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B3").FormulaArrayA1 = "C1:D3*2";

            Assert.Throws<InvalidOperationException>(() => ws.Row(2).Delete());
            Assert.Throws<InvalidOperationException>(() => ws.Column(1).Delete());
            Assert.Throws<InvalidOperationException>(() => ws.Rows("1,3").Delete());

            AssertArrayFormula(ws, "A1:B3", "C1:D3*2");
        }

        [TestCase("row delete")]
        [TestCase("cell insert above")]
        [TestCase("cell insert before")]
        [TestCase("column part delete")]
        [Issue("102")]
        public void Change_of_part_of_a_data_table_throws_when_only_the_first_cell_has_the_formula(string operation)
        {
            // A data table from a file has the formula only in its first cell.
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet();
            var formula = XLCellFormula.DataTable1D(Area.Parse("B2:C4"), new Point(1, 1), false, true);
            ws.Internals.CellsCollection.FormulaSlice.Set(new Point(2, 2), formula);

            Action change = operation switch
            {
                "row delete" => () => ws.Row(3).Delete(),
                "cell insert above" => () => ws.Cell("C3").InsertCellsAbove(1),
                "cell insert before" => () => ws.Cell("C3").InsertCellsBefore(1),
                _ => () => ws.Range("C2:C4").Delete(XLShiftDeletedCells.ShiftCellsLeft),
            };
            Assert.Throws<InvalidOperationException>(() => change());

            Assert.AreSame(formula, ws.Cell("B2").Formula);
            Assert.AreEqual("B2:C4", formula.Range.ToString());
        }

        [Test]
        [Issue("102")]
        public void Insert_of_zero_rows_or_columns_throws_an_argument_exception()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B2").FormulaArrayA1 = "C1:D2*2";

            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Row(1).InsertRowsBelow(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Column(1).InsertColumnsAfter(0));

            AssertArrayFormula(ws, "A1:B2", "C1:D2*2");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_all_rows_of_an_array_formula_deletes_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A2:B4").FormulaArrayA1 = "C2:D4*2";
            ws.Cell("A5").Value = 5;

            ws.Rows(2, 4).Delete();

            Assert.False(ws.Cell("A2").HasFormula);
            Assert.False(ws.Cell("B4").HasFormula);
            Assert.AreEqual(5, ws.Cell("A2").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_all_columns_of_an_array_formula_deletes_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B1:C2").FormulaArrayA1 = "E1:F2*2";
            ws.Cell("D1").Value = 5;

            ws.Columns(2, 3).Delete();

            Assert.False(ws.Cell("B1").HasFormula);
            Assert.False(ws.Cell("C2").HasFormula);
            Assert.AreEqual(5, ws.Cell("B1").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_range_rows_that_contain_all_rows_of_an_array_formula_deletes_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A2:A3").FormulaArrayA1 = "B2:B3*2";
            ws.Cell("A4").Value = 4;

            ws.Range("A1:C5").Rows(2, 3).Delete();

            Assert.False(ws.Cell("A3").HasFormula);
            Assert.AreEqual(4, ws.Cell("A2").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_range_columns_that_contain_all_columns_of_an_array_formula_deletes_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B1:C1").FormulaArrayA1 = "B2:C2*2";
            ws.Cell("D1").Value = 4;

            ws.Range("A1:E3").Columns(2, 3).Delete();

            Assert.False(ws.Cell("C1").HasFormula);
            Assert.AreEqual(4, ws.Cell("B1").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_range_rows_with_different_columns_throws_when_an_array_formula_can_move()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A2:A3").FormulaArrayA1 = "B2:B3*2";
            var rows = ws.Range("A2:A2").Rows();
            rows.Add(ws.Range("A3:B3").FirstRow());

            Assert.Throws<InvalidOperationException>(() => rows.Delete());

            AssertArrayFormula(ws, "A2:A3", "B2:B3*2");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_range_rows_with_different_columns_below_an_array_formula_deletes_the_rows()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B1").FormulaArrayA1 = "C1:D1*2";
            ws.Cell("A4").Value = 4;
            var rows = ws.Range("A2:A2").Rows();
            rows.Add(ws.Range("A3:B3").FirstRow());

            rows.Delete();

            AssertArrayFormula(ws, "A1:B1", "C1:D1*2");
            Assert.AreEqual(4, ws.Cell("A2").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_range_rows_with_different_columns_keeps_an_array_formula_in_other_columns()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B5:B6").FormulaArrayA1 = "D5:D6*2";
            ws.Cell("A3").Value = 3;
            ws.Cell("C4").Value = 4;
            var rows = ws.Range("A2:A2").Rows();
            rows.Add(ws.Range("C3:C3").FirstRow());

            rows.Delete();

            AssertArrayFormula(ws, "B5:B6", "D5:D6*2");
            Assert.AreEqual(3, ws.Cell("A2").Value);
            Assert.AreEqual(4, ws.Cell("C3").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_the_same_row_two_times_throws_before_the_first_delete()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A3:A4").FormulaArrayA1 = "B3:B4*2";

            // The first delete moves the array formula up. Then the second delete splits it.
            Assert.Throws<InvalidOperationException>(() => ws.Rows("2,2").Delete());

            AssertArrayFormula(ws, "A3:A4", "B3:B4*2");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_the_first_row_of_an_array_formula_two_times_deletes_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A3:A4").FormulaArrayA1 = "B3:B4*2";
            ws.Cell("A5").Value = 5;

            // The second delete removes the row that moved up after the first delete.
            ws.Rows("3,3").Delete();

            Assert.False(ws.Cell("A3").HasFormula);
            Assert.AreEqual(5, ws.Cell("A3").Value);
        }

        [TestCase(true)]
        [TestCase(false)]
        [Issue("102")]
        public void Delete_of_rows_of_two_sheets_checks_all_sheets_before_the_first_delete(bool allRowsOfArray)
        {
            using var wb = new XLWorkbook();
            var ws1 = wb.AddWorksheet();
            var ws2 = wb.AddWorksheet();
            ws1.Range("A3:A4").FormulaArrayA1 = "B3:B4*2";
            ws2.Cell("A1").Value = 1;

            // The collection deletes the rows of the second sheet first.
            var rows = new XLRows(wb, null);
            rows.Add((XLRow)ws2.Row(1));
            rows.Add((XLRow)ws1.Row(3));
            if (allRowsOfArray)
                rows.Add((XLRow)ws1.Row(4));

            if (allRowsOfArray)
            {
                rows.Delete();
                Assert.False(ws1.Cell("A3").HasFormula);
                Assert.True(ws2.Cell("A1").IsEmpty());
            }
            else
            {
                Assert.Throws<InvalidOperationException>(() => rows.Delete());
                AssertArrayFormula(ws1, "A3:A4", "B3:B4*2");
                Assert.AreEqual(1, ws2.Cell("A1").Value);
            }
        }

        [Test]
        [Issue("102")]
        public void Delete_of_range_rows_ignores_a_range_row_that_an_earlier_delete_removed()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:B1").FormulaArrayA1 = "C1:D1*2";
            ws.Cell("A6").Value = 6;
            var rows = ws.Range("A1:C5").Rows(3, 5);
            ws.Row(4).Delete();

            rows.Delete();

            AssertArrayFormula(ws, "A1:B1", "C1:D1*2");
            Assert.AreEqual(6, ws.Cell("A3").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_table_rows_that_moves_part_of_an_array_formula_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { new[] { "First", "Second" }, new[] { "a", "b" }, new[] { "c", "d" } });
            var table = ws.Range("A1:B3").CreateTable();
            ws.Range("A5:C5").FormulaArrayA1 = "1+1";

            Assert.Throws<InvalidOperationException>(() => table.DataRange.Rows().Delete());

            AssertArrayFormula(ws, "A5:C5", "1+1");
            Assert.AreEqual("c", ws.Cell("A3").Value);
        }

        [Test]
        [Issue("102")]
        public void Delete_of_table_rows_above_an_array_formula_moves_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { new[] { "First", "Second" }, new[] { "a", "b" }, new[] { "c", "d" } });
            var table = ws.Range("A1:B3").CreateTable();
            ws.Range("A5:B6").FormulaArrayA1 = "1+1";

            table.DataRange.Rows().Delete();

            AssertArrayFormula(ws, "A3:B4", "1+1");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_overlapping_columns_throws_before_the_first_delete()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("C1:D1").FormulaArrayA1 = "C2:D2*2";

            Assert.Throws<InvalidOperationException>(() => ws.Columns("A:B,B").Delete());

            AssertArrayFormula(ws, "C1:D1", "C2:D2*2");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_a_table_column_that_splits_an_array_formula_throws_and_keeps_the_table()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { new[] { "First", "Second", "Third" }, new[] { "a", "b", "c" } });
            var table = ws.Range("A1:C2").CreateTable();
            ws.Range("E2:E3").FormulaArrayA1 = "F2:F3*2";

            Assert.Throws<InvalidOperationException>(() => table.Column(2).Delete());

            CollectionAssert.AreEqual(new[] { "First", "Second", "Third" }, table.Fields.Select(f => f.Name));
            Assert.AreEqual(1, table.Field("Second").Index);
            AssertArrayFormula(ws, "E2:E3", "F2:F3*2");
        }

        [Test]
        [Issue("102")]
        public void Insert_that_moves_part_of_an_array_formula_out_of_the_sheet_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1048574:A1048575").FormulaArrayA1 = "1";

            Assert.Throws<InvalidOperationException>(() => ws.Row(1).InsertRowsAbove(2));
            ws.Row(1).InsertRowsAbove(1);

            AssertArrayFormula(ws, "A1048575:A1048576", "1");
        }

        [Test]
        [Issue("102")]
        public void Delete_of_a_row_above_an_array_formula_moves_the_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A2:B3").FormulaArrayA1 = "C2:D3*2";

            ws.Row(1).Delete();
            ws.Range("C5:D5").Delete(XLShiftDeletedCells.ShiftCellsUp);
            ws.Range("A5:B5").Delete(XLShiftDeletedCells.ShiftCellsLeft);

            AssertArrayFormula(ws, "A1:B2", "C1:D2*2");
        }

        [Test]
        [Issue("110")]
        public void Transpose_of_a_column_moves_an_array_formula_to_the_right_of_the_range()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Range("E1:E3").FormulaArrayA1 = "1+1";

            ws.Range("A1:A3").Transpose(XLTransposeOptions.MoveCells);

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("C1").Value);
            AssertArrayFormula(ws, "G1:G3", "1+1");
        }

        [Test]
        [Issue("110")]
        public void Transpose_throws_before_the_insert_when_the_delete_splits_an_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Cell("B1").Value = "right";
            ws.Range("C10:D10").FormulaArrayA1 = "1+1";

            Assert.Throws<InvalidOperationException>(() => ws.Range("A1:A3").Transpose(XLTransposeOptions.MoveCells));

            Assert.AreEqual(2, ws.Cell("A2").Value);
            Assert.AreEqual("right", ws.Cell("B1").Value);
            AssertArrayFormula(ws, "C10:D10", "1+1");
        }

        [Test]
        [Issue("110")]
        public void Transpose_of_a_range_with_an_array_formula_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1:A2").FormulaArrayA1 = "1+1";
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 2;

            Assert.Throws<InvalidOperationException>(() => ws.Range("A1:B2").Transpose(XLTransposeOptions.MoveCells));

            AssertArrayFormula(ws, "A1:A2", "1+1");
            Assert.AreEqual(1, ws.Cell("B1").Value);
            Assert.AreEqual(2, ws.Cell("B2").Value);
        }

        [Test]
        [Issue("110")]
        public void Transpose_keeps_an_array_formula_in_one_cell_on_the_diagonal()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("A1").FormulaArrayA1 = "1+1";
            ws.Cell("B1").Value = 1;
            ws.Cell("A2").Value = 2;
            ws.Cell("B2").Value = 3;

            ws.Range("A1:B2").Transpose(XLTransposeOptions.MoveCells);

            AssertArrayFormula(ws, "A1:A1", "1+1");
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(1, ws.Cell("A2").Value);
        }

        [Test]
        [Issue("110")]
        public void Transpose_keeps_an_array_formula_that_has_only_a_diagonal_cell_in_the_square()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("C3:D4").FormulaArrayA1 = "1+1";
            ws.Cell("B1").Value = 1;
            ws.Cell("A2").Value = 2;

            ws.Range("A1:C3").Transpose(XLTransposeOptions.MoveCells);

            AssertArrayFormula(ws, "C3:D4", "1+1");
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(1, ws.Cell("A2").Value);
        }

        [Test]
        [Issue("110")]
        public void Transpose_of_one_cell_of_an_array_formula_keeps_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:C3").FormulaArrayA1 = "1+1";

            ws.Range("B2").Transpose(XLTransposeOptions.MoveCells);

            AssertArrayFormula(ws, "B2:C3", "1+1");
        }

        [Test]
        [Issue("110")]
        public void Transpose_with_a_clear_throws_when_the_clear_removes_part_of_an_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Range("C1:D1").FormulaArrayA1 = "1+1";

            Assert.Throws<InvalidOperationException>(() => ws.Range("A1:A3").Transpose(XLTransposeOptions.ReplaceCells));

            Assert.AreEqual(2, ws.Cell("A2").Value);
            AssertArrayFormula(ws, "C1:D1", "1+1");
        }

        [Test]
        [Issue("110")]
        public void Transpose_with_a_clear_removes_an_array_formula_in_the_cleared_cells()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Range("B1:C1").FormulaArrayA1 = "1+1";

            ws.Range("A1:A3").Transpose(XLTransposeOptions.ReplaceCells);

            Assert.AreEqual(1, ws.Cell("A1").Value);
            Assert.AreEqual(2, ws.Cell("B1").Value);
            Assert.AreEqual(3, ws.Cell("C1").Value);
            Assert.False(ws.Cell("B1").HasFormula);
            Assert.False(ws.Cell("C1").HasFormula);
        }

        [Test]
        [Issue("110")]
        public void Resize_of_a_table_throws_before_a_change_when_the_delete_of_the_totals_row_splits_an_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var table = CreateTableWithTotalsRow(ws);
            ws.Range("B5:C5").FormulaArrayA1 = "1+1";
            ws.Cell("A7").Value = "below";

            Assert.Throws<InvalidOperationException>(() => table.Resize("A1:C6"));

            Assert.AreEqual("A1:B4", table.RangeAddress.ToString());
            Assert.True(ws.Cell("C1").IsEmpty());
            Assert.AreEqual("below", ws.Cell("A7").Value);
            AssertArrayFormula(ws, "B5:C5", "1+1");
        }

        [Test]
        [Issue("110")]
        public void Resize_of_a_table_to_more_rows_keeps_an_array_formula_below_the_table()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var table = CreateTableWithTotalsRow(ws);
            ws.Range("A9:B9").FormulaArrayA1 = "1+1";

            table.Resize("A1:B6");

            Assert.AreEqual("A1:B6", table.RangeAddress.ToString());
            AssertArrayFormula(ws, "A9:B9", "1+1");
        }

        [Test]
        [Issue("110")]
        public void Resize_of_a_table_to_fewer_rows_throws_before_a_change_when_the_insert_splits_an_array_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var table = CreateTableWithTotalsRow(ws);
            ws.Range("C3:D3").FormulaArrayA1 = "1+1";

            Assert.Throws<InvalidOperationException>(() => table.Resize("A1:C3"));

            Assert.AreEqual("A1:B4", table.RangeAddress.ToString());
            Assert.True(ws.Cell("C1").IsEmpty());
            AssertArrayFormula(ws, "C3:D3", "1+1");
        }

        private static IXLTable CreateTableWithTotalsRow(IXLWorksheet ws)
        {
            ws.Cell("A1").InsertData(new[] { new[] { "First", "Second" }, new[] { "a", "b" }, new[] { "c", "d" } });
            var table = ws.Range("A1:B3").CreateTable();
            table.ShowTotalsRow = true;
            return table;
        }

        private static void AssertArrayFormula(IXLWorksheet ws, string range, string formula)
        {
            foreach (var cell in ws.Range(range).Cells())
            {
                Assert.True(cell.HasArrayFormula, cell.Address.ToString());
                Assert.AreEqual(range, cell.FormulaReference!.ToString());
                Assert.AreEqual(formula, cell.FormulaA1);
            }
        }

        [Test]
        [Issue("64")]
        public void Formula_from_null_has_empty_value()
        {
            var formula = new XLFormula((string)null);

            Assert.AreEqual(string.Empty, formula.Value);
            Assert.IsFalse(formula.IsFormula);
        }

        [Test]
        [Issue("64")]
        public void Formula_without_value_has_empty_value()
        {
            var formula = new XLFormula();

            Assert.AreEqual(string.Empty, formula.Value);
        }
    }
}
