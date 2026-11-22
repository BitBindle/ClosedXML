using System;
using System.Globalization;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Extensions
{
    [TestFixture]
    public class ObjectExtensionsTests
    {
        // A value that parses back at 15 significant digits keeps its text. Excel shows 15 digits.
        [TestCase(0d, "0")]
        [TestCase(0.1d, "0.1")]
        [TestCase(-2.5d, "-2.5")]
        [TestCase(123456789012345d, "123456789012345")]
        [TestCase(1e-300d, "1E-300")]
        [TestCase(1.5e300d, "1.5E+300")]
        [TestCase(double.MaxValue, "1.7976931348623157E+308")]
        [TestCase(double.Epsilon, "4.94065645841247E-324")]
        // A value that needs 16 or 17 significant digits gets them, so the text parses back to the same value.
        [TestCase(0.1d + 0.2d, "0.30000000000000004")]
        [TestCase(1234567.123456789d, "1234567.123456789")]
        [TestCase(9007199254740992d, "9007199254740992")]
        [TestCase(1234567890123456d, "1234567890123456")]
        [TestCase(1.0000000000000002d, "1.0000000000000002")]
        [Issue("14")]
        public void Double_text_has_fewest_digits_that_parse_back_to_same_value(double value, string expectedText)
        {
            var text = ObjectExtensions.ToRoundTripString(value);

            Assert.AreEqual(expectedText, text);
            Assert.AreEqual(value, double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
            Assert.AreEqual(expectedText, value.ToInvariantString());
        }

        [TestCase(-0d, "-0")]
        [TestCase(1d, "1")]
        [TestCase(-42d, "-42")]
        [TestCase(999999999999999d, "999999999999999")]
        [TestCase(1e15, "1E+15")]
        [TestCase(1e16, "1E+16")]
        [TestCase(9007199254740993d, "9007199254740992")]
        [TestCase(12345678901234567890d, "1.2345678901234567E+19")]
        [TestCase(0.0001d, "0.0001")]
        [TestCase(0.00012345d, "0.00012345")]
        [TestCase(1e-5, "1E-05")]
        [TestCase(1.5e-7, "1.5E-07")]
        [TestCase(double.MinValue, "-1.7976931348623157E+308")]
        [TestCase(2.2250738585072014E-308, "2.2250738585072014E-308")] // The smallest normal value.
        [TestCase(2.2250738585072009E-308, "2.225073858507201E-308")] // The largest subnormal value.
        [TestCase(0.99999999999999989d, "0.9999999999999999")]
        [TestCase(-0.86136133483801225d, "-0.8613613348380122")]
        // The exact value is in the middle of two 16-digit texts. The text has the even digit.
        [TestCase(562949953421312.25d, "562949953421312.2")]
        [TestCase(562949953421312.75d, "562949953421312.8")]
        // The rounding adds a digit: 9.9999999999999995E+22 becomes 1E+23.
        [TestCase(1e23, "1E+23")]
        [Issue("34")]
        public void Double_text_has_correctly_rounded_digits(double value, string expectedText)
        {
            var text = value.ToInvariantString();

            Assert.AreEqual(expectedText, text);
            Assert.AreEqual(BitConverter.DoubleToInt64Bits(value), BitConverter.DoubleToInt64Bits(double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)));
        }

        [TestCase(double.NaN, "NaN")]
        [TestCase(double.PositiveInfinity, "Infinity")]
        [TestCase(double.NegativeInfinity, "-Infinity")]
        [Issue("14")]
        public void Double_special_values_keep_their_text(double value, string expectedText)
        {
            Assert.AreEqual(expectedText, ObjectExtensions.ToRoundTripString(value));
        }

        // A value that parses back at 7 significant digits keeps its text.
        [TestCase(0f, "0")]
        [TestCase(0.1f, "0.1")]
        [TestCase(1234567f, "1234567")]
        // A value that needs 8 or 9 significant digits gets them, so the text parses back to the same value.
        [TestCase(16777216f, "16777216")]
        [TestCase(123456.79f, "123456.79")]
        [TestCase(0.3f + 0.6f, "0.90000004")]
        [TestCase(float.MaxValue, "3.4028235E+38")]
        [TestCase(1000.00006f, "1000.00006")]
        [TestCase(1000.00024f, "1000.00024")]
        [Issue("14")]
        public void Float_text_has_fewest_digits_that_parse_back_to_same_value(float value, string expectedText)
        {
            var text = ObjectExtensions.ToRoundTripString(value);

            Assert.AreEqual(expectedText, text);
            Assert.AreEqual(value, float.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
            Assert.AreEqual(expectedText, value.ToInvariantString());
        }

        [TestCase(-0f, "-0")]
        [TestCase(1e7f, "1E+07")]
        [TestCase(1e-5f, "1E-05")]
        [TestCase(float.Epsilon, "1.401298E-45")]
        [Issue("34")]
        public void Float_text_has_correctly_rounded_digits(float value, string expectedText)
        {
            var text = value.ToInvariantString();

            Assert.AreEqual(expectedText, text);
            Assert.AreEqual(BitConverter.SingleToInt32Bits(value), BitConverter.SingleToInt32Bits(float.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture)));
        }

        [TestCase(float.NaN, "NaN")]
        [TestCase(float.PositiveInfinity, "Infinity")]
        [TestCase(float.NegativeInfinity, "-Infinity")]
        [Issue("14")]
        public void Float_special_values_keep_their_text(float value, string expectedText)
        {
            Assert.AreEqual(expectedText, ObjectExtensions.ToRoundTripString(value));
        }

        [Test]
        [Issue("60")]
        public void Invariant_text_of_other_struct_is_ToString_text()
        {
            // A struct without a special format gives the text of its ToString method.
            var guid = new Guid("0f8fad5b-d9cb-469f-a165-70867728950e");

            Assert.AreEqual("0f8fad5b-d9cb-469f-a165-70867728950e", guid.ToInvariantString());
        }
    }
}
