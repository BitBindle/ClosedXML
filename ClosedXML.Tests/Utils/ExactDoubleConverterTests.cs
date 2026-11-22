using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Utils;
using NUnit.Framework;

namespace ClosedXML.Tests.Utils;

/// <summary>
/// The expected values are the IEEE 754 results. Python and .NET Core 3.0 and later give the same results.
/// The tests also run on .NET Framework, where the parser and the formatter of the runtime are not exact.
/// </summary>
[TestFixture]
internal class ExactDoubleConverterTests
{
    /// <summary>
    /// Number cells of a workbook that Excel saved with 17 significant digits, and the text that ClosedXML
    /// writes with the fewest digits from 15 to 17. The workbook is the input of the external reference sparkline test.
    /// </summary>
    private static readonly TestCaseData[] ExcelNumbers =
    {
        new TestCaseData("42.043900000000001", 0x4045059E83E425AFUL, "42.0439"),
        new TestCaseData("1.8654637783131349E-5", 0x3EF38F90F278D788UL, "1.865463778313135E-05"),
        new TestCaseData("8.5000000000000006E-2", 0x3FB5C28F5C28F5C3UL, "0.085"),
        new TestCaseData("7.0974999999999996E-2", 0x3FB22B6AE7D566CFUL, "0.070975"),
        new TestCaseData("5.8427534995500012E-2", 0x3FADEA36BFFBAECFUL, "0.05842753499550001"),
        new TestCaseData("7.1730669903414104E-2", 0x3FB25CF0F15AF73DUL, "0.0717306699034141"),
        new TestCaseData("68.738276465294106", 0x40512F3FEBEE760FUL, "68.7382764652941"),
        new TestCaseData("0.64949595641478497", 0x3FE4C8ABBE75F19BUL, "0.649495956414785"),
        new TestCaseData("1176.4705882352939", 0x409261E1E1E1E1E1UL, "1176.470588235294"),
        new TestCaseData("11.576328541025971", 0x4027271488D6EA77UL, "11.57632854102597"),
        new TestCaseData("16.076328541025969", 0x4030138A446B753BUL, "16.07632854102597"),
        new TestCaseData("0.31961180303614128", 0x3FD47485105D2C29UL, "0.3196118030361413"),
        new TestCaseData("4.8196118030361417", 0x401347485105D2C3UL, "4.819611803036142"),
        new TestCaseData("4.4999999999999998E-2", 0x3FA70A3D70A3D70AUL, "0.045"),
        new TestCaseData("3.2799597532906931E-3", 0x3F6ADE92FBEC2C61UL, "0.003279959753290693"),
        new TestCaseData("8.6394960319995793", 0x4021476C061EB84DUL, "8.63949603199958"),
        new TestCaseData("-0.10346701834730029", 0xBFBA7CD084042B5EUL, "-0.1034670183473003"),
        new TestCaseData("1.3759264999163581E-6", 0x3EB7158F2F850001UL, "1.375926499916358E-06"),
        new TestCaseData("-0.86136133483801225", 0xBFEB9045A56561A2UL, "-0.8613613348380122"),
        new TestCaseData("-3.5342534968048879E-2", 0xBFA2186AAFB0E635UL, "-0.03534253496804888"),
        new TestCaseData("-1.255439488689891E-5", 0xBEEA5416E50CFFFFUL, "-1.255439488689891E-05"),
        new TestCaseData("0.91836691248611024", 0x3FED634301DB671EUL, "0.9183669124861102"),
        new TestCaseData("8.1633087513889813E-2", 0x3FB4E5E7F124C714UL, "0.08163308751388981"),
        new TestCaseData("71.923671458974027", 0x4051FB1D6EE522B1UL, "71.92367145897403"),
        new TestCaseData("5.8332631425899999", 0x40175542EEE98B81UL, "5.83326314259"),
        new TestCaseData("5.7165978797381998", 0x4016DDCBD5A76F1DUL, "5.7165978797382"),
        new TestCaseData("117.64705882352941", 0x405D696969696969UL, "117.6470588235294"),
        new TestCaseData("0.21418159555235439", 0x3FCB6A4D7226B70CUL, "0.2141815955523544"),
        new TestCaseData("251.97834770865211", 0x406F7F4E9FDA98E0UL, "251.9783477086521"),
        new TestCaseData("2.3553612818440761E-5", 0x3EF8B29FF2B85BD7UL, "2.355361281844076E-05"),
        new TestCaseData("1.0256220499444959", 0x3FF068F2AAA916E4UL, "1.025622049944496"),
        new TestCaseData("1.4214999981278845E-3", 0x3F574A340034CF9BUL, "0.0014214999981278845"),
        new TestCaseData("89648.105656933447", 0x40F5E301B0C5531CUL, "89648.10565693345"),
        new TestCaseData("1.3123206336756597E-2", 0x3F8AE056F047BD92UL, "0.013123206336756597"),
    };

    [TestCaseSource(nameof(ExcelNumbers))]
    [Issue("34")]
    public void TryParse_reads_nearest_double_of_excel_number(string excelText, ulong expectedBits, string _)
    {
        Assert.True(ExactDoubleConverter.TryParse(excelText, out var value));
        Assert.AreEqual(expectedBits, (ulong)BitConverter.DoubleToInt64Bits(value));
    }

    [TestCaseSource(nameof(ExcelNumbers))]
    [Issue("34")]
    public void Format_writes_fewest_digits_from_15_to_17_of_excel_number(string _, ulong bits, string expectedText)
    {
        var value = BitConverter.Int64BitsToDouble((long)bits);

        Assert.AreEqual(expectedText, ExactDoubleConverter.Format(value));
    }

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
        Assert.True(ExactDoubleConverter.TryParse(text, out var value));
        Assert.AreEqual(expectedBits, (ulong)BitConverter.DoubleToInt64Bits(value));
    }

    [Test]
    public void TryParse_uses_all_digits_of_long_text()
    {
        // The parser keeps 800 significant digits. A non-zero digit after them decides a tie.
        var tie = "9007199254740993." + new string('0', 900);
        var aboveTie = tie + "1";
        var longInteger = new string('1', 850) + "e-845";

        Assert.True(ExactDoubleConverter.TryParse(tie, out var tieValue));
        Assert.True(ExactDoubleConverter.TryParse(aboveTie, out var aboveTieValue));
        Assert.True(ExactDoubleConverter.TryParse(longInteger, out var longIntegerValue));
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
        Assert.False(ExactDoubleConverter.TryParse(text, out _));
    }

    [TestCase("1.8e308")]
    [TestCase("1e400")]
    [TestCase("1e999999")]
    [TestCase("NaN")]
    public void TryParse_gives_result_of_runtime_outside_exact_range(string text)
    {
        // An overflow and the symbols of the culture have the result of the runtime.
        var expected = double.TryParse(text, XLHelper.NumberStyle, CultureInfo.InvariantCulture, out var expectedValue);

        var actual = ExactDoubleConverter.TryParse(text, out var value);

        Assert.AreEqual(expected, actual);
        Assert.AreEqual(expectedValue, value);
    }

    [TestCase(0x3FD3333333333334UL, "0.30000000000000004")]
    [TestCase(0x3FEB0E7009B61CE0UL, "0.8455124082255701")]
    [TestCase(0x7FEFFFFFFFFFFFFFUL, "1.7976931348623157E+308")]
    [TestCase(0x0000000000000001UL, "4.94065645841247E-324")]
    [TestCase(0x000FFFFFFFFFFFFFUL, "2.225073858507201E-308")]
    [TestCase(0x44B52D02C7E14AF6UL, "1E+23")]
    [TestCase(0x430BB8629B743EDEUL, "975319753197531.8")]
    public void Format_writes_fewest_digits_from_15_to_17_of_double(ulong bits, string expectedText)
    {
        var value = BitConverter.Int64BitsToDouble((long)bits);

        Assert.AreEqual(expectedText, ExactDoubleConverter.Format(value));
    }

    [TestCase(0d, "0")]
    [TestCase(-0d, "-0")]
    [TestCase(123d, "123")]
    [TestCase(-1.5d, "-1.5")]
    [TestCase(123.456d, "123.456")]
    [TestCase(999999999999999d, "999999999999999")]
    [TestCase(1e15d, "1E+15")]
    [TestCase(1234567890123455d, "1234567890123455")]
    [TestCase(1e19d, "1E+19")]
    [TestCase(0.0001d, "0.0001")]
    [TestCase(0.00001d, "1E-05")]
    [TestCase(double.NaN, "NaN")]
    [TestCase(double.PositiveInfinity, "Infinity")]
    [TestCase(double.NegativeInfinity, "-Infinity")]
    public void Format_uses_notation_of_general_format(double value, string expectedText)
    {
        Assert.AreEqual(expectedText, ExactDoubleConverter.Format(value));
    }

#if NET
    // .NET Core 3.0 and later format and parse exactly, so the runtime gives the expected results.

    [Test]
    public void Format_gives_same_text_as_runtime()
    {
        foreach (var value in GetRandomDoubles(20000))
            Assert.AreEqual(FormatWithRuntime(value), ExactDoubleConverter.Format(value), "Bits {0:X16}", BitConverter.DoubleToInt64Bits(value));
    }

    [Test]
    public void TryParse_gives_same_value_as_runtime()
    {
        var random = new Random(34);
        for (var i = 0; i < 20000; i++)
        {
            var digits = new string(Enumerable.Range(0, random.Next(1, 40)).Select(_ => (char)('0' + random.Next(10))).ToArray());
            var pointPosition = random.Next(digits.Length + 1);
            var text = digits.Insert(pointPosition, ".") + "e" + random.Next(-345, 330).ToString(CultureInfo.InvariantCulture);
            var expected = double.Parse(text, CultureInfo.InvariantCulture);

            Assert.True(ExactDoubleConverter.TryParse(text, out var value), text);
            Assert.AreEqual(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(value), text);
        }
    }

    [Test]
    public void TryParse_gives_same_value_as_runtime_for_text_that_excel_writes()
    {
        // Excel writes 17 significant digits. The integer of 16 to 19 digits is larger than 2^53 in most cases.
        var random = new Random(34);
        for (var i = 0; i < 20000; i++)
        {
            var value = random.NextDouble() * Math.Pow(10, random.Next(-12, 19));
            var digitCount = random.Next(16, 20);
            var text = value.ToString("E" + (digitCount - 1).ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            var expected = double.Parse(text, CultureInfo.InvariantCulture);

            Assert.True(ExactDoubleConverter.TryParse(text, out var parsed), text);
            Assert.AreEqual(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(parsed), text);
        }
    }

    private static IEnumerable<double> GetRandomDoubles(int count)
    {
        var random = new Random(34);
        var bytes = new byte[8];
        for (var i = 0; i < count; i++)
        {
            random.NextBytes(bytes);
            var value = BitConverter.ToDouble(bytes, 0);
            if (!double.IsNaN(value) && !double.IsInfinity(value))
                yield return value;

            // Values with few digits, as a person types them.
            yield return Math.Round(random.NextDouble() * Math.Pow(10, random.Next(-8, 18)), random.Next(0, 10));

            // Values with all digits, as a calculation gives them.
            yield return random.NextDouble() * Math.Pow(10, random.Next(-13, 19));
        }
    }

    private static string FormatWithRuntime(double value)
    {
        foreach (var format in new[] { "G15", "G16" })
        {
            var text = value.ToString(format, CultureInfo.InvariantCulture);
            if (double.Parse(text, CultureInfo.InvariantCulture).Equals(value))
                return text;
        }

        return value.ToString("G17", CultureInfo.InvariantCulture);
    }
#endif
}
