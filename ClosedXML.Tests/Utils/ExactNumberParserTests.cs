using System;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Utils;
using NUnit.Framework;

namespace ClosedXML.Tests.Utils;

/// <summary>
/// The expected values are the IEEE 754 results. Python and .NET Core 3.0 and later give the same results.
/// The tests also run on .NET Framework, where the parser of the runtime is not exact.
/// </summary>
[TestFixture]
internal class ExactNumberParserTests
{
    [TestCase("0.84551240822557", 0x3FEB0E7009B61CDFUL)]
    [TestCase("0.84551240822557006", 0x3FEB0E7009B61CE0UL)]
    [TestCase("0.1", 0x3FB999999999999AUL)]
    [TestCase("1e23", 0x44B52D02C7E14AF6UL)]
    [TestCase("123456789012345678901234567890", 0x45F8EE90FF6C373EUL)]
    [TestCase("9007199254740993", 0x4340000000000000UL)]
    [TestCase("9007199254740993.0000000000000000001", 0x4340000000000001UL)]
    [TestCase("9007199254740991.5", 0x4340000000000000UL)]
    [TestCase("9007199254740995", 0x4340000000000002UL)]
    [TestCase("4503599627370496.5", 0x4330000000000000UL)]
    [TestCase("4503599627370497.5", 0x4330000000000002UL)]
    [TestCase("9007199254740993e5", 0x44486A0000000001UL)]
    [TestCase("1234567890123456789e-27", 0x3E1535AFDF5AE86EUL)]
    [TestCase("1234567890123456789e27", 0x494BAE1279C66306UL)]
    [TestCase("9999999999999999999e-27", 0x3E45798EE2308C3AUL)]
    [TestCase("1.7976931348623157E+308", 0x7FEFFFFFFFFFFFFFUL)]
    [TestCase("2.2250738585072014E-308", 0x0010000000000000UL)]
    [TestCase("2.2250738585072011E-308", 0x000FFFFFFFFFFFFFUL)]
    [TestCase("4.9406564584124654E-324", 0x0000000000000001UL)]
    [TestCase("2.4703282292062328E-324", 0x0000000000000001UL)]
    [TestCase("2.4703282292062327E-324", 0x0000000000000000UL)]
    [TestCase("1e-400", 0x0000000000000000UL)]
    [TestCase("1e-999999", 0x0000000000000000UL)]
    [TestCase("-0", 0x8000000000000000UL)]
    [TestCase("0.000", 0x0000000000000000UL)]
    [TestCase(" \t+1.5E+3\r\n", 0x4097700000000000UL)]
    [TestCase(".5", 0x3FE0000000000000UL)]
    [TestCase("5.", 0x4014000000000000UL)]
    [TestCase("-00012.50e-1", 0xBFF4000000000000UL)]
    public void TryParse_reads_nearest_double(string text, ulong expectedBits)
    {
        Assert.True(ExactNumberParser.TryParse(text, out var value));
        Assert.AreEqual(expectedBits, (ulong)BitConverter.DoubleToInt64Bits(value));
    }

    [Test]
    public void TryParse_uses_all_digits_of_long_text()
    {
        // The parser keeps 800 significant digits. A non-zero digit after them decides a tie.
        var tie = "9007199254740993." + new string('0', 900);
        var aboveTie = tie + "1";
        var longInteger = new string('1', 850) + "e-845";

        Assert.True(ExactNumberParser.TryParse(tie, out var tieValue));
        Assert.True(ExactNumberParser.TryParse(aboveTie, out var aboveTieValue));
        Assert.True(ExactNumberParser.TryParse(longInteger, out var longIntegerValue));
        Assert.AreEqual(9007199254740992d, tieValue);
        Assert.AreEqual(9007199254740994d, aboveTieValue);
        Assert.AreEqual(0x40C5B38E38E38E39UL, (ulong)BitConverter.DoubleToInt64Bits(longIntegerValue));
    }

    [TestCase("")]
    [TestCase("abc")]
    [TestCase(".")]
    [TestCase("1.2.3")]
    [TestCase("1e")]
    [TestCase("1e+")]
    [TestCase("e5")]
    [TestCase("1,5")]
    [TestCase("1.5x")]
    [TestCase("- 1")]
    public void TryParse_rejects_text_that_is_not_number(string text)
    {
        Assert.False(ExactNumberParser.TryParse(text, out _));
    }

    [TestCase("1.8e308")]
    [TestCase("1e400")]
    [TestCase("1e999999")]
    [TestCase("NaN")]
    public void TryParse_gives_result_of_runtime_outside_exact_range(string text)
    {
        // An overflow and the symbols of the culture have the result of the runtime.
        var expected = double.TryParse(text, XLHelper.NumberStyle, CultureInfo.InvariantCulture, out var expectedValue);

        var actual = ExactNumberParser.TryParse(text, out var value);

        Assert.AreEqual(expected, actual);
        Assert.AreEqual(expectedValue, value);
    }

#if NET
    // .NET Core 3.0 and later parse exactly, so the runtime gives the expected results.

    [Test]
    public void TryParse_gives_same_value_as_runtime()
    {
        var random = new Random(12345);
        for (var i = 0; i < 20000; i++)
        {
            var digits = new string(Enumerable.Range(0, random.Next(1, 40)).Select(_ => (char)('0' + random.Next(10))).ToArray());
            var pointPosition = random.Next(digits.Length + 1);
            var text = digits.Insert(pointPosition, ".") + "e" + random.Next(-345, 330).ToString(CultureInfo.InvariantCulture);
            var expected = double.Parse(text, CultureInfo.InvariantCulture);

            Assert.True(ExactNumberParser.TryParse(text, out var value), text);
            Assert.AreEqual(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(value), text);
        }
    }

    [Test]
    public void TryParse_gives_same_value_as_runtime_for_text_that_excel_writes()
    {
        // Excel writes 17 significant digits. The integer of 16 to 19 digits is larger than 2^53 in most cases.
        var random = new Random(12345);
        for (var i = 0; i < 20000; i++)
        {
            var value = random.NextDouble() * Math.Pow(10, random.Next(-12, 19));
            var digitCount = random.Next(16, 20);
            var text = value.ToString("E" + (digitCount - 1).ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            var expected = double.Parse(text, CultureInfo.InvariantCulture);

            Assert.True(ExactNumberParser.TryParse(text, out var parsed), text);
            Assert.AreEqual(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(parsed), text);
        }
    }
#endif
}
