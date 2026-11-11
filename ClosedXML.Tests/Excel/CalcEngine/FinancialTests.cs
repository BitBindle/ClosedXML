using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class FinancialTests
    {
        [TestCase("FV(0.06/12,10,-200,-500,1)", 2581.4033740601362)]
        [TestCase("FV(0.12/12,12,-1000)", 12682.503013196976)]
        [TestCase("FV(0.11/12,35,-2000,,1)", 82846.24637190059)]
        [TestCase("FV(0.06/12,12,-100,-1000,1)", 2301.4018303409139)]
        public void Fv_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [TestCase("FV(0,1,1000)", -1000)] // Zero interest rate
        [TestCase("FV(0,5,10000,5000)", -55000.00)] // Zero interest rate with present value
        [TestCase("FV(-0.4,2,1000)", -1600.00)] // Negative interest rate
        [TestCase("FV(0.01,0.5,1000)", -498.75621120889502)] // Non-integer period
        [TestCase("FV(0.1,-2,1000)", 1735.5371900826453)] // Negative periods
        [TestCase("FV(0.1,2,0,4)", -4.84)] // No PMT, but present value
        [TestCase("FV(0,2,-1000)", 2000.00)] // Negative PMT - money is paid to us
        [TestCase("FV(0.000001,1000,1000)", -1000499.6661261424)] // Small number and high number of periods, check for stability
        public void Fv_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [Test]
        public void Fv_DefaultFutureValueIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("FV(0.1,2,1000)"), XLWorkbook.EvaluateExpr("FV(0.1,2,1000,0)"));
        }

        [Test]
        public void Fv_DefaultTypeIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("FV(0.1,5,1000)"), XLWorkbook.EvaluateExpr("FV(0.1,5,1000,0,0)"));
        }

        [Test]
        public void Fv_ZeroPeriodsReturnsPresentValue()
        {
            Assert.AreEqual(-100, XLWorkbook.EvaluateExpr("FV(0.1,0,1000, 100)"));
        }

        [TestCase("IPMT(0.1/12,1,3*12,8000)", -66.666666666666686)]
        [TestCase("IPMT(0.1,3,3,8000)", -292.4471299093658)]
        public void Ipmt_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [TestCase("IPMT(0,1,1,1000)", 0)] // Zero interest rate
        [TestCase("IPMT(0,1,5,10000,5000)", 0)] // Zero interest rate with future value
        [TestCase("IPMT(-0.4,1,2,1000)", 400.00)] // Negative interest rate
        [TestCase("IPMT(0.01,1,0.5,1000)", -10.00)] // Non-integer period
        [TestCase("IPMT(0.01,1,1.4,1000)", -10.00)] // Different non-integer period
        [TestCase("IPMT(0.1,1,2,0,4)", 0)] // No principal, but future value
        [TestCase("IPMT(0.1,1,2,-1000)", 100)] // Negative principal - money is paid to us
        [TestCase("IPMT(0.000001,1,1000,1000)", -0.001)] // Small number and high number of periods, check for stability
        public void Ipmt_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [Test]
        public void Ipmt_DefaultFutureValueIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("IPMT(0.1,1,2,1000)"), XLWorkbook.EvaluateExpr("IPMT(0.1,1,2,1000,0)"));
        }

        [Test]
        public void Ipmt_DefaultTypeIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("IPMT(0.1,1,5,1000)"), XLWorkbook.EvaluateExpr("IPMT(0.1,1,5,1000,0,0)"));
        }

        [Test]
        public void Ipmt_ZeroOrNegativePeriodsReturnsNumError()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("IPMT(0.1,1,0,1000)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("IPMT(0.1,1,-1,1000)"));
        }

        [TestCase(-1)]
        [TestCase(-1.5)]
        [TestCase(-100)]
        public void Ipmt_RateLessOrEqualMinusOneReturnsNumError(double rate)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"IPMT({rate},2,3,1000,10000,1)"));
        }

        [Test]
        public void Ipmt_PeriodOutOfRangeReturnsNumError()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("IPMT(0.1,0,1,1000)"));
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("IPMT(0.1,2,1,1000)"));
        }

        [TestCase("IRR(A2:A6)", -0.021244848273410961)]
        [TestCase("IRR(A2:A7)", 0.086630948036531622)]
        [TestCase("IRR(A2:A4,-10%)", -0.44350694133474067)]
        public void Irr_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[] { "Data", -70000, 12000, 15000, 18000, 21000, 26000 });

            var actual = (double)ws.Evaluate(formula);
            Assert.AreEqual(expectedResult, actual, 1e-9);
        }

        [Test]
        public void Irr_IgnoresTextLogicalAndBlankInReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[] { -100, "text", true, Blank.Value, 110 });

            Assert.AreEqual(0.1, (double)ws.Evaluate("IRR(A1:A5)"), 1e-9);
        }

        [TestCase("IRR({100,200})")] // No negative value
        [TestCase("IRR({-100,-200})")] // No positive value
        public void Irr_WithoutResultReturnsNumError(string formula)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr(formula));
        }

        [Test]
        public void Irr_PropagatesErrorInValues()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("IRR({-100,#DIV/0!,110})"));
        }

        [TestCase("NPER(0.12/12,-100,-1000,10000,1)", 59.673865674294571)]
        [TestCase("NPER(0.12/12,-100,-1000,10000)", 60.082122853761661)]
        [TestCase("NPER(0.12/12,-100,-1000)", -9.5785940398131613)]
        public void Nper_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [TestCase("NPER(0,-100,1000)", 10)] // Zero interest rate
        [TestCase("NPER(0,-100,1000,500)", 15)] // Zero interest rate with future value
        public void Nper_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [TestCase("NPER(0,0,1000)")] // Zero interest rate and no payment
        [TestCase("NPER(0.1,-50,1000)")] // Payment is less than interest, logarithm of negative number
        [TestCase("NPER(-1,-100,1000)")] // Rate less or equal -1
        public void Nper_WithoutResultReturnsNumError(string formula)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr(formula));
        }

        [TestCase("NPV(10%,-10000,3000,4200,6800)", 1188.4434123352216)]
        [TestCase("NPV(8%,8000,9200,10000,12000,14500)-40000", 1922.0615549323628)]
        [TestCase("NPV(8%,8000,9200,10000,12000,14500,-9000)-40000", -3749.4650870155206)]
        public void Npv_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [Test]
        public void Npv_ReferenceOnlyCountsNumbers()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new object[] { 110, "text", true, Blank.Value, 121 });

            // Values in the reference that are not numbers are skipped and don't take a period.
            Assert.AreEqual(200, (double)ws.Evaluate("NPV(10%,A1:A5)"), XLHelper.Epsilon);
        }

        [Test]
        public void Npv_PropagatesErrorInReference()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 110;
            ws.Cell("A2").FormulaA1 = "1/0";

            Assert.AreEqual(XLError.DivisionByZero, ws.Evaluate("NPV(10%,A1:A2)"));
        }

        [Test]
        public void Npv_ScalarArgumentsAreConvertedToNumbers()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("NPV(10%,1,110)"), XLWorkbook.EvaluateExpr("NPV(10%,TRUE,\"110\")"));
            Assert.AreEqual(XLError.IncompatibleValue, XLWorkbook.EvaluateExpr("NPV(10%,\"text\")"));
        }

        [Test]
        public void Npv_RateMinusOneReturnsDivisionByZeroError()
        {
            Assert.AreEqual(XLError.DivisionByZero, XLWorkbook.EvaluateExpr("NPV(-1,100)"));
        }

        [TestCase("PMT(0.08/12,10,10000)", -1037.03208935915)]
        [TestCase("PMT(0.08/12,10,10000,0,1)", -1030.16432717797)]
        public void Pmt_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [Test]
        public void Pmt_PaymentsMustPayForPrincipalAndFutureValue()
        {
            var actual = (double)XLWorkbook.EvaluateExpr("PMT(0,2,5000,10000)");
            Assert.AreEqual(-7500, actual);
        }

        [TestCase("PMT(0,1,1000)", -1000)] // Zero interest rate
        [TestCase("PMT(0,5,10000,5000)", -3000)] // Zero interest rate for 5 years, (10k principal, pay all and have 5k in bank at the end = payment is 3k/year)
        [TestCase("PMT(-0.4,2,1000)", -225)] // Negative interest rate
        [TestCase("PMT(0.01,0.5,1000)", -2014.98756211209)] // Non-integer period
        [TestCase("PMT(0.1,-2,1000)", 476.19047619048)] // Negative periods
        [TestCase("PMT(0.1,2,0,4)", -1.90476190476)] // No principal, but future value
        [TestCase("PMT(0,2,-1000)", 500)] // Negative principal - money is paid to us
        [TestCase("PMT(0.000001,1000,1000)", -1.00050058333321)] // Small number and high number of periods, check for stability
        public void Pmt_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [Test]
        public void Pmt_TypeConvertsNumberToOneOrZero()
        {
            // Spec says "if type is any number other than 0 or 1, #NUM! is returned.", but Excel accepts any number as type
            var formulaFormat = "PMT(0.1,2,1000,500,{0})";
            var zeroType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "0"));
            var oneType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "1"));
            var nonZeroType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "0.000001"));

            Assert.AreNotEqual(zeroType, oneType);
            Assert.AreEqual(oneType, nonZeroType);
        }

        [Test]
        public void Pmt_DefaultFutureValueIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("PMT(0.1,2,1000)"), XLWorkbook.EvaluateExpr("PMT(0.1,2,1000,0)"));
        }

        [Test]
        public void Pmt_DefaultTypeIsZero()
        {
            Assert.AreEqual(XLWorkbook.EvaluateExpr("PMT(0.1,5,1000)"), XLWorkbook.EvaluateExpr("PMT(0.1,5,1000,0,0)"));
        }

        [Test]
        public void Pmt_ZeroPeriodsReturnsNumError()
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr("PMT(0.1,0,1000)"));
        }

        [TestCase(-1)]
        [TestCase(-1.5)]
        [TestCase(-100)]
        public void Pmt_RateLessOrEqualMinusOneReturnsNumError(double rate)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"PMT({rate},1,1000,5000,1)"));
        }

        [TestCase("PV(0.08/12,12*20,500,,0)", -59777.145851187769)]
        public void Pv_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [TestCase("PV(0,10,-100)", 1000)] // Zero interest rate
        [TestCase("PV(0,10,-100,-500)", 1500)] // Zero interest rate with future value
        [TestCase("PV(0.1,2,0,-121)", 100)] // No payment, only future value
        [TestCase("PV(0.1,2,-110,0,1)", 210)] // Payments at the beginning of the period
        public void Pv_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, XLHelper.Epsilon);
        }

        [TestCase(-1)]
        [TestCase(-2)]
        public void Pv_RateLessOrEqualMinusOneReturnsNumError(double rate)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr($"PV({rate},2,100)"));
        }

        [TestCase("RATE(4*12,-200,8000)", 0.0077014724882333765)]
        [TestCase("RATE(4*12,-200,8000)*12", 0.092417669858800518)]
        public void Rate_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, 1e-9);
        }

        [TestCase("RATE(2,0,-100,121)", 0.1)] // No payment, only future value
        [TestCase("RATE(10,-100,1000)", 0)] // Zero interest rate
        [TestCase("RATE(2,-110,210,0,1)", 0.1)] // Payments at the beginning of the period
        [TestCase("RATE(2,0,-100,121,0,0.5)", 0.1)] // Different guess
        public void Rate_EdgeCases(string formula, double expectedResult)
        {
            var actual = (double)XLWorkbook.EvaluateExpr(formula);
            Assert.AreEqual(expectedResult, actual, 1e-9);
        }

        [TestCase("RATE(0,-100,1000)")] // No period
        [TestCase("RATE(10,100,1000)")] // No solution, all cash flows are positive
        public void Rate_WithoutResultReturnsNumError(string formula)
        {
            Assert.AreEqual(XLError.NumberInvalid, XLWorkbook.EvaluateExpr(formula));
        }
    }
}
