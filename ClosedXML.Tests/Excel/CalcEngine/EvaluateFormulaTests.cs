using System;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine.Exceptions;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// <see cref="IXLWorksheet.EvaluateFormula"/> gives the result of a formula without implicit intersection. A caller
    /// needs it for formulas that give a list, for example the source of a data validation list, a conditional format
    /// formula or a defined name.
    /// </summary>
    [TestFixture]
    public class EvaluateFormulaTests
    {
        [Test]
        public void Scalar_result_is_a_value()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var result = ws.EvaluateFormula("1+2");

            Assert.AreEqual(XLFormulaResultType.Value, result.Type);
            Assert.AreEqual(XLCellValue.FromObject(3), result.Value);
        }

        [Test]
        public void Error_result_is_a_value()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var result = ws.EvaluateFormula("1/0");

            Assert.AreEqual(XLFormulaResultType.Value, result.Type);
            Assert.AreEqual(XLCellValue.FromObject(XLError.DivisionByZero), result.Value);
        }

        [Test]
        public void Range_of_a_list_source_gives_all_cells_of_the_range()
        {
            // The source of a data validation list, as Excel saves it.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "Red";
            ws.Cell("A2").Value = "Green";
            ws.Cell("A3").Value = "Blue";

            var result = ws.EvaluateFormula("$A$1:$A$3", "C2");

            Assert.AreEqual(XLFormulaResultType.Reference, result.Type);
            var range = result.Ranges.Single();
            Assert.AreSame(ws, range.Worksheet);
            Assert.AreEqual("A1:A3", range.RangeAddress.ToStringRelative());
            Assert.AreEqual(new[] { "Red", "Green", "Blue" }, range.Cells().Select(c => c.GetText()));
        }

        [Test]
        public void Evaluate_gives_one_cell_of_the_same_range()
        {
            // Evaluate applies implicit intersection, so the caller cannot get the other cells.
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = "Red";
            ws.Cell("A2").Value = "Green";
            ws.Cell("A3").Value = "Blue";

            Assert.AreEqual(XLCellValue.FromObject("Green"), ws.Evaluate("$A$1:$A$3", "C2"));
            Assert.AreEqual(XLCellValue.FromObject(XLError.IncompatibleValue), ws.Evaluate("$A$1:$A$3", "C5"));
        }

        [Test]
        public void Function_that_gives_a_reference_gives_the_reference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");
            ws.Cell("A1").Value = 3;

            var result = ws.EvaluateFormula("OFFSET($B$1,0,0,$A$1,1)", "D1");

            Assert.AreEqual(XLFormulaResultType.Reference, result.Type);
            Assert.AreEqual("B1:B3", result.Ranges.Single().RangeAddress.ToStringRelative());
        }

        [Test]
        public void Union_gives_each_area_in_the_order_of_the_formula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet");

            var result = ws.EvaluateFormula("($C$1:$C$2,$A$1,$C$1:$C$2)");

            Assert.AreEqual(XLFormulaResultType.Reference, result.Type);
            Assert.AreEqual(new[] { "C1:C2", "A1:A1", "C1:C2" }, result.Ranges.Select(r => r.RangeAddress.ToStringRelative()));
            Assert.That(result.Ranges.All(r => r.Worksheet == ws));
        }

        [Test]
        public void Reference_to_another_sheet_has_the_other_sheet()
        {
            using var wb = new XLWorkbook();
            var form = wb.AddWorksheet("Form");
            var data = wb.AddWorksheet("Data");
            data.Cell("B2").Value = "Item";

            var result = form.EvaluateFormula("Data!$B$1:$B$2", "A1");

            var range = result.Ranges.Single();
            Assert.AreSame(data, range.Worksheet);
            Assert.AreEqual("B1:B2", range.RangeAddress.ToStringRelative());
            Assert.AreEqual(XLCellValue.FromObject("Item"), range.LastCell().Value);
        }

        [Test]
        public void Reference_to_a_deleted_sheet_is_an_error()
        {
            using var wb = new XLWorkbook();
            var form = wb.AddWorksheet("Form");
            wb.AddWorksheet("Data");
            wb.Worksheet("Data").Delete();

            var result = form.EvaluateFormula("Data!$A$1:$A$3", "B1");

            Assert.AreEqual(XLFormulaResultType.Value, result.Type);
            Assert.AreEqual(XLCellValue.FromObject(XLError.CellReference), result.Value);
        }

        [Test]
        public void Reference_to_a_sheet_span_gives_one_range_for_each_sheet()
        {
            using var wb = new XLWorkbook();
            var form = wb.AddWorksheet("Form");
            var middle = wb.AddWorksheet("Middle");
            var data = wb.AddWorksheet("Data");

            var result = form.EvaluateFormula("Form:Data!$A$1", "B1");

            Assert.AreEqual(XLFormulaResultType.Reference, result.Type);
            Assert.AreEqual(new[] { form, middle, data }, result.Ranges.Select(r => r.Worksheet));
            Assert.That(result.Ranges.All(r => r.RangeAddress.ToStringRelative() == "A1:A1"));

            // Evaluate cannot intersect a reference to more than one sheet. Without an address, it has no cell to intersect.
            Assert.AreEqual(XLCellValue.FromObject(XLError.IncompatibleValue), form.Evaluate("Form:Data!$A$1", "B1"));
            Assert.Throws<MissingContextException>(() => form.Evaluate("Form:Data!$A$1"));
        }

        [Test]
        public void Defined_name_gives_its_reference()
        {
            using var wb = new XLWorkbook();
            var form = wb.AddWorksheet("Form");
            var data = wb.AddWorksheet("Data");
            wb.DefinedNames.Add("Items", "Data!$A$1:$A$4");

            var result = form.EvaluateFormula("Items");

            var range = result.Ranges.Single();
            Assert.AreSame(data, range.Worksheet);
            Assert.AreEqual("A1:A4", range.RangeAddress.ToStringRelative());
        }

        [Test]
        public void Array_constant_gives_all_elements()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var result = ws.EvaluateFormula("{1,\"a\";TRUE,#N/A}");

            Assert.AreEqual(XLFormulaResultType.Array, result.Type);
            var expected = new XLCellValue[,]
            {
                { 1, "a" },
                { true, XLError.NoValueAvailable },
            };
            Assert.AreEqual(expected, result.Array);
        }

        [Test]
        public void Array_expression_gives_all_elements()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var result = ws.EvaluateFormula("{1,2,3}*10", "A1");

            Assert.AreEqual(XLFormulaResultType.Array, result.Type);
            Assert.AreEqual(new XLCellValue[,] { { 10, 20, 30 } }, result.Array);
            Assert.AreEqual(XLCellValue.FromObject(10), ws.Evaluate("{1,2,3}*10", "A1"));
        }

        [Test]
        public void Array_keeps_a_blank_element()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 1;

            var result = ws.EvaluateFormula("TRANSPOSE($A$1:$A$2)", "C1");

            Assert.AreEqual(new XLCellValue[,] { { 1, Blank.Value } }, result.Array);
        }

        [Test]
        public void Dirty_cell_is_calculated_first()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "2*3";

            var result = ws.EvaluateFormula("A1*2");

            Assert.AreEqual(XLCellValue.FromObject(12), result.Value);
        }

        [Test]
        public void Property_of_another_type_throws()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var value = ws.EvaluateFormula("1");
            var reference = ws.EvaluateFormula("A1");
            var array = ws.EvaluateFormula("{1,2}");

            Assert.Throws<InvalidOperationException>(() => _ = value.Ranges);
            Assert.Throws<InvalidOperationException>(() => _ = value.Array);
            Assert.Throws<InvalidOperationException>(() => _ = reference.Value);
            Assert.Throws<InvalidOperationException>(() => _ = array.Value);
        }
    }
}
