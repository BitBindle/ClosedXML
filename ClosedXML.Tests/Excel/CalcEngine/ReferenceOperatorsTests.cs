using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class ReferenceOperatorsTests
    {
        #region Implicit intersection

        [Test]
        public void ImplicitIntersection_DoesNotAffectSingleCellReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D5").FormulaA1 = "ABS(B3:B3)";

            Assert.AreEqual(1, ws.Cell("D5").Value);
        }

        [Test]
        public void ImplicitIntersection_TakesReferenceFromHorizontalLine()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D3").FormulaA1 = "ABS(B1:B10)";

            Assert.AreEqual(1, ws.Cell("D3").Value);
        }

        [Test]
        public void ImplicitIntersection_TakesReferenceFromVerticalLine()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("B5").FormulaA1 = "ABS(A3:Z3)";

            Assert.AreEqual(1, ws.Cell("B5").Value);
        }

        [Test]
        public void ImplicitIntersection_TakesReferenceEvenFromIntersectionEvenFromDifferentSheet()
        {
            using var wb = new XLWorkbook();
            var sheet1 = wb.AddWorksheet("Sheet1");
            sheet1.Cell("B3").Value = -1;

            var sheet2 = wb.AddWorksheet("Sheet2");
            sheet2.Cell("D3").FormulaA1 = "ABS(Sheet1!B1:B10)";

            Assert.AreEqual(1, sheet2.Cell("D3").Value);
        }

        [Test]
        public void ImplicitIntersection_WithoutIntersectionResultsInValueError()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D5").FormulaA1 = "ABS(B1:B4)";

            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("D5").Value);
        }

        [Test]
        public void ImplicitIntersection_CanWorkOnlyWithOneArea()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            ws.Cell("D3").FormulaA1 = "ABS((B1:B2,B3:B5))"; // A continous range made of two areas

            Assert.AreEqual(XLError.IncompatibleValue, ws.Cell("D3").Value);
        }

        [Test]
        public void ImplicitIntersection_IntersectionMustHaveSpanOfOneCell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("B3").Value = -1;
            var horizontalIntersectionCell = ws.Cell("D3");
            horizontalIntersectionCell.FormulaA1 = "ABS(A1:B5)";
            Assert.AreEqual(XLError.IncompatibleValue, horizontalIntersectionCell.Value);

            var verticalIntersectionCell = ws.Cell("B5");
            verticalIntersectionCell.FormulaA1 = "ABS(A3:C4)";
            Assert.AreEqual(XLError.IncompatibleValue, verticalIntersectionCell.Value);
        }

        [TestCase("B2", "A1:A3*10", 20)]
        [TestCase("B5", "A1:A3*10", XLError.IncompatibleValue)]
        [TestCase("B5", "A1:B3*10", XLError.IncompatibleValue)]
        [TestCase("A5", "-A1:C1", -1)]
        [TestCase("D5", "-A1:C1", XLError.IncompatibleValue)]
        [TestCase("D2", "A1:A3&\"x\"", "2x")]
        [TestCase("D3", "A1:A3=3", true)]
        [TestCase("D2", "SUM(A1:A3*10)", 20)]
        [TestCase("D5", "SUM(A1:A3*10)", XLError.IncompatibleValue)]
        [TestCase("D2", "{1,2,3}*10", 10)]
        public void ImplicitIntersection_OperatorTakesOneCellOfRangeOperand(string address, string formula, object expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Cell(address).FormulaA1 = formula;

            Assert.AreEqual(expected, ws.Cell(address).Value);
        }

        [TestCase("SUMPRODUCT(A1:A3*B1:B3)", 32)]
        [TestCase("SUMPRODUCT((A1:A3>1)*B1:B3)", 11)]
        [TestCase("MMULT(TRANSPOSE(A1:A3*1),B1:B3)", 32)]
        [TestCase("ROWS(A1:A3*1)", 3)]
        [TestCase("SUMPRODUCT(ABS(-A1:A3))", 6)]
        [TestCase("SUMPRODUCT(INDEX(A1:B3,0,2))", 15)]
        [TestCase("SUMPRODUCT(COUNTIF(A1:A3,A1:A3))", 3)]
        [TestCase("SUMPRODUCT(SUMIF(A1:A3,A1:A3,B1:B3))", 15)]
        [TestCase("INDEX(A1:A3*10,2)", 20)]
        [TestCase("MATCH(1,INDEX((A1:A3=2)*(B1:B3=5),0,1),0)", 2)]
        public void ImplicitIntersection_ArgumentOfArrayParameterIsEvaluatedAsArray(string formula, double expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { new[] { 1, 4 }, new[] { 2, 5 }, new[] { 3, 6 } });
            ws.Cell("D5").FormulaA1 = formula;

            Assert.AreEqual(expected, ws.Cell("D5").Value);
        }

        [Test]
        public void ImplicitIntersection_ArrayFormulaDoesNotIntersectOperands()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Range("B5:B7").FormulaArrayA1 = "A1:A3*10";

            Assert.AreEqual(10, ws.Cell("B5").Value);
            Assert.AreEqual(20, ws.Cell("B6").Value);
            Assert.AreEqual(30, ws.Cell("B7").Value);
        }

        [TestCase("B2", "@A1:A3", 2)]
        [TestCase("B5", "@A1:A3", XLError.IncompatibleValue)]
        [TestCase("B2", "@A1:B3", XLError.IncompatibleValue)]
        [TestCase("B2", "@(A1,A2)", XLError.IncompatibleValue)]
        [TestCase("B5", "@A3", 3)]
        [TestCase("B2", "10*@A1:A3", 20)]
        [TestCase("B2", "@INDEX({4,5,6},0)", 4)]
        [TestCase("B3", "_xlfn.SINGLE(A1:A3)", 3)]
        [TestCase("B3", "SINGLE({7,8})", 7)]
        public void ImplicitIntersectionOperator_TakesOneValue(string address, string formula, object expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Cell(address).FormulaA1 = formula;

            Assert.AreEqual(expected, ws.Cell(address).Value);
        }

        #endregion

        #region Reference range operator

        [TestCase("A1:B2", 4)]
        [TestCase("A1:B5:C3", 3 * 5)]
        [TestCase("A1:C3:B5", 3 * 5)]
        [TestCase("A1:C3:B2", 3 * 3)]
        [TestCase("Sheet1!A1:B2", 4)]
        [TestCase("Sheet1!A1:Sheet1!B2", 4)]
        [TestCase("Sheet1!A1:Sheet1!B2", 4)]
        [TestCase("A1:Sheet1!B2", 4)]
        [TestCase("Sheet1!B2:C5:Sheet1!D3", 12)]
        [TestCase("(Sheet1!A1,A5):B5", 10)]
        [TestCase("B5:(Sheet1!A1,A5)", 10)]
        public void Range_UnifiesReferencesIntoSingleAreas(string referenceFormula, int expectedCellCount)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cells("A1:Z100").Value = 1;

            var referenceCells = ws.Evaluate($"SUM({referenceFormula})");
            Assert.AreEqual(expectedCellCount, referenceCells);
        }

        [TestCase("Sheet1!A1:C5")]
        [TestCase("Sheet1!A1:B3:C5")]
        [TestCase("Sheet1!A1:B3:C4:Sheet1!B5:C5")]
        public void Range_LeftSideDeterminesSheetIfRightOmitted(string formula)
        {
            using var wb = new XLWorkbook();
            var firstSheet = wb.AddWorksheet("Sheet1");
            firstSheet.Cells("A1:C5").Value = 1;
            var secondSheet = wb.AddWorksheet("Sheet2");
            secondSheet.Cell("A1").FormulaA1 = $"=SUM({formula})";

            Assert.AreEqual(15, secondSheet.Cell("A1").Value);
        }

        [TestCase("Current!A1:Other!B2")]
        [TestCase("A1:Other!B2")]
        [TestCase("A1:(Other!B2,C3)")]
        [TestCase("Other!A1:(Other!B2,C3)")] // C3 is taken from current worksheet since multiple areas on rhs
        [TestCase("(Other!A1,A5):Other!B2")] // A5 is taken from current worksheet since multiple areas on lhs
        [TestCase("(Current!A1):Other!B2")]
        // [TestCase("Other!A5:(B5)")] This causes #VALUE! in Excel, but it shouldn't. It's likely there is a "Fast parser for simple sheet areas" and "Full path" for complicated operands and they behave inconsistenly
        public void Range_UnificationAcrossSheetsResultsInValueError(string referenceFormula)
        {
            using var wb = new XLWorkbook();
            var formulaSheet = wb.AddWorksheet("Current");
            wb.AddWorksheet("Other");

            Assert.AreEqual(XLError.IncompatibleValue, formulaSheet.Evaluate($"SUM({referenceFormula})"));
        }

        [TestCase("A1:IF(TRUE,1,)")]
        [TestCase("IF(TRUE,1,):A1")]
        [TestCase("IF(TRUE,\"text\"):A1")]
        [TestCase("IF(TRUE,FALSE):A1")]
        public void Range_OnlyReferencesCanBeRange(string referenceFormula)
        {
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();

            Assert.AreEqual(XLError.IncompatibleValue, sheet.Evaluate($"SUM({referenceFormula})"));
        }

        #endregion

        #region Reference union

        [TestCase("A1,A2", 2)]
        [TestCase("A1:A3,B1", 4)]
        [TestCase("A1,B1:B3", 4)]
        [TestCase("Other!A1,Current!A1", 11)]
        [TestCase("A1,Other!A1", 11)]
        [TestCase("B2:D3,B2:D3", 12)] // Full overlap
        [TestCase("A1:B3,B1:C3", 12)] // Partial overlap
        [TestCase("Current!A1:B3,Other!B1:C3", 66)]
        [TestCase("A1,Other!A1,Current!A1", 10 + 1 + 1)]
        [TestCase("A1:B2,Other!A1:B2,B2:C3,Other!E5:Other!F6", 4 + 40 + 4 + 40)]
        public void Union_CanJoinAnyTwoRanges(string formula, int expectedSum)
        {
            using var wb = new XLWorkbook();
            var currentSheet = wb.AddWorksheet("Current");
            currentSheet.Cells("A1:F10").Value = 1;
            var otherSheet = wb.AddWorksheet("Other");
            otherSheet.Cells("A1:F10").Value = 10;

            // Not extra braces, so the comma is interpreted as union and not an extra argument
            var value = currentSheet.Evaluate($"SUM(({formula}))");

            Assert.AreEqual(expectedSum, value);
        }

        #endregion
    }
}
