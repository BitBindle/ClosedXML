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

        // .NET Framework formats these values with a different last digit.
        [TestCase(-0.86136133483801225d, "-0.8613613348380122")]
        [TestCase(562949953421312.25d, "562949953421312.2")]
        [Issue("34")]
        public void Double_text_is_the_same_on_each_runtime(double value, string expectedText)
        {
            Assert.AreEqual(expectedText, value.ToInvariantString());
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

        [TestCase(float.NaN, "NaN")]
        [TestCase(float.PositiveInfinity, "Infinity")]
        [TestCase(float.NegativeInfinity, "-Infinity")]
        [Issue("14")]
        public void Float_special_values_keep_their_text(float value, string expectedText)
        {
            Assert.AreEqual(expectedText, ObjectExtensions.ToRoundTripString(value));
        }
    }
}
