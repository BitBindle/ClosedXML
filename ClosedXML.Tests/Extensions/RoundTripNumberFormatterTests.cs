using System;
using System.Globalization;
using ClosedXML.Utils;
using NUnit.Framework;

namespace ClosedXML.Tests.Extensions
{
    [TestFixture]
    public class RoundTripNumberFormatterTests
    {
        // The expected texts are the texts of .NET Core. The tests also run on .NET Framework,
        // where the formatter and the parser of the runtime can give a different last digit.
        [TestCase(0d, "0")]
        [TestCase(-0d, "-0")]
        [TestCase(1d, "1")]
        [TestCase(-42d, "-42")]
        [TestCase(999999999999999d, "999999999999999")]
        [TestCase(1e15, "1E+15")]
        [TestCase(1e16, "1E+16")]
        [TestCase(1234567890123456d, "1234567890123456")]
        [TestCase(9007199254740993d, "9007199254740992")]
        [TestCase(12345678901234567890d, "1.2345678901234567E+19")]
        [TestCase(0.1d, "0.1")]
        [TestCase(0.1d + 0.2d, "0.30000000000000004")]
        [TestCase(0.0001d, "0.0001")]
        [TestCase(0.00012345d, "0.00012345")]
        [TestCase(1e-5, "1E-05")]
        [TestCase(1.5e-7, "1.5E-07")]
        [TestCase(1.5e300, "1.5E+300")]
        [TestCase(double.MaxValue, "1.7976931348623157E+308")]
        [TestCase(double.MinValue, "-1.7976931348623157E+308")]
        [TestCase(double.Epsilon, "4.94065645841247E-324")]
        [TestCase(2.2250738585072014E-308, "2.2250738585072014E-308")] // The smallest normal value.
        [TestCase(2.2250738585072009E-308, "2.225073858507201E-308")] // The largest subnormal value.
        [TestCase(1.0000000000000002d, "1.0000000000000002")]
        [TestCase(0.99999999999999989d, "0.9999999999999999")]
        // .NET Framework writes -0.8613613348380123 for this value.
        [TestCase(-0.86136133483801225d, "-0.8613613348380122")]
        // The exact value is in the middle of two 16-digit texts. The text has the even digit.
        [TestCase(562949953421312.25d, "562949953421312.2")]
        [TestCase(562949953421312.75d, "562949953421312.8")]
        // The rounding adds a digit: 9.9999999999999995E+22 becomes 1E+23.
        [TestCase(1e23, "1E+23")]
        [Issue("34")]
        public void Double_text_is_the_same_as_on_net_core(double value, string expectedText)
        {
            var text = RoundTripNumberFormatter.Format(value);

            Assert.AreEqual(expectedText, text);
        }

        [TestCase(double.NaN, "NaN")]
        [TestCase(double.PositiveInfinity, "Infinity")]
        [TestCase(double.NegativeInfinity, "-Infinity")]
        [Issue("34")]
        public void Double_special_values_keep_their_text(double value, string expectedText)
        {
            Assert.AreEqual(expectedText, RoundTripNumberFormatter.Format(value));
        }

        [TestCase(0f, "0")]
        [TestCase(-0f, "-0")]
        [TestCase(0.1f, "0.1")]
        [TestCase(1234567f, "1234567")]
        [TestCase(1e7f, "1E+07")]
        [TestCase(16777216f, "16777216")]
        [TestCase(123456.79f, "123456.79")]
        [TestCase(0.3f + 0.6f, "0.90000004")]
        [TestCase(1000.00006f, "1000.00006")]
        [TestCase(1e-5f, "1E-05")]
        [TestCase(float.MaxValue, "3.4028235E+38")]
        [TestCase(float.Epsilon, "1.401298E-45")]
        [Issue("34")]
        public void Float_text_is_the_same_as_on_net_core(float value, string expectedText)
        {
            var text = RoundTripNumberFormatter.Format(value);

            Assert.AreEqual(expectedText, text);
        }

        [TestCase(float.NaN, "NaN")]
        [TestCase(float.PositiveInfinity, "Infinity")]
        [TestCase(float.NegativeInfinity, "-Infinity")]
        [Issue("34")]
        public void Float_special_values_keep_their_text(float value, string expectedText)
        {
            Assert.AreEqual(expectedText, RoundTripNumberFormatter.Format(value));
        }

#if NETCOREAPP
        // The formatter and the parser of .NET Core 3.0 and later are exact. Thus they are the
        // reference for the formatter. .NET Framework is not a reference, thus this test is only
        // for .NET Core.
        [Test]
        [Issue("34")]
        public void Text_of_random_values_is_the_same_as_on_net_core()
        {
            var random = new Random(34);
            var bytes = new byte[8];
            for (var i = 0; i < 20000; i++)
            {
                random.NextBytes(bytes);
                var doubleValue = BitConverter.ToDouble(bytes, 0);
                if (double.IsFinite(doubleValue))
                    Assert.AreEqual(NetCoreText(doubleValue), RoundTripNumberFormatter.Format(doubleValue), "Bits {0:X16}", BitConverter.DoubleToInt64Bits(doubleValue));

                var scaledValue = random.NextDouble() * Math.Pow(10, random.Next(-8, 16));
                Assert.AreEqual(NetCoreText(scaledValue), RoundTripNumberFormatter.Format(scaledValue), "Bits {0:X16}", BitConverter.DoubleToInt64Bits(scaledValue));

                var floatValue = BitConverter.ToSingle(bytes, 0);
                if (float.IsFinite(floatValue))
                    Assert.AreEqual(NetCoreText(floatValue), RoundTripNumberFormatter.Format(floatValue), "Bits {0:X8}", BitConverter.SingleToInt32Bits(floatValue));
            }
        }

        private static string NetCoreText(double value)
        {
            foreach (var format in new[] { "G15", "G16" })
            {
                var text = value.ToString(format, CultureInfo.InvariantCulture);
                if (double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture).Equals(value))
                    return text;
            }

            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        private static string NetCoreText(float value)
        {
            foreach (var format in new[] { "G7", "G8" })
            {
                var text = value.ToString(format, CultureInfo.InvariantCulture);
                if (float.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture).Equals(value))
                    return text;
            }

            return value.ToString("G9", CultureInfo.InvariantCulture);
        }
#endif
    }
}
