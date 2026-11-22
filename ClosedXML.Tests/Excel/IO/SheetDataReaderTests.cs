using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;
using StringItem = ClosedXML.Excel.CalcEngine.OneOf<string, ClosedXML.Excel.XLImmutableRichText>;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class SheetDataReaderTests
{
    private const string Ns = OpenXmlConst.Main2006SsNs;

    [Test]
    public void Empty_sheetData_is_ok()
    {
        AssertSheetData(
            "<sheetData/>",
            (_, ws) => Assert.True(ws.Internals.CellsCollection.IsEmpty));
    }

    [Test]
    public void Tracks_row_number_when_r_is_omitted()
    {
        AssertSheetData(
            """
            <sheetData>
              <row>
                <c r="A1"><v>1</v></c>
              </row>
              <row>
                <c r="A2"><v>2</v></c>
              </row>
              <row r="5">
                <c r="A5"><v>5</v></c>
              </row>
              <row>
                <c r="A6"><v>6</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(1, ws.Cell("A1").GetDouble());
                Assert.AreEqual(2, ws.Cell("A2").GetDouble());
                Assert.AreEqual(5, ws.Cell("A5").GetDouble());
                Assert.AreEqual(6, ws.Cell("A6").GetDouble());
            });
    }

    [Test]
    public void Tracks_column_when_cell_r_is_omitted()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c><v>1</v></c>
                <c><v>2</v></c>
                <c r="D1"><v>4</v></c>
                <c><v>5</v></c>
              </row>
              <row r="2">
                <c><v>10</v></c>
                <c><v>11</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(1, ws.Cell("A1").GetDouble());
                Assert.AreEqual(2, ws.Cell("B1").GetDouble());
                Assert.AreEqual(4, ws.Cell("D1").GetDouble());
                Assert.AreEqual(5, ws.Cell("E1").GetDouble());
                Assert.AreEqual(10, ws.Cell("A2").GetDouble());
                Assert.AreEqual(11, ws.Cell("B2").GetDouble());
            });
    }

    [Test]
    public void Default_cell_type_is_number()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1"><v>1.5</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(XLDataType.Number, ws.Cell("A1").DataType);
                Assert.AreEqual(1.5, ws.Cell("A1").GetDouble());
            });
    }

    [Test]
    public void Number_without_value_stays_blank()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="n"/>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.True(ws.Cell("A1").Value.IsBlank));
    }

    [Test]
    public void Unparsable_number_stays_blank()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="n"><v>not-a-number</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.True(ws.Cell("A1").Value.IsBlank));
    }

    /// <summary>
    /// Number cells of a workbook that Excel saved with 17 significant digits, and the bits of the nearest double.
    /// </summary>
    private static readonly TestCaseData[] ExcelNumbers =
    {
        new TestCaseData("42.043900000000001", 0x4045059E83E425AFUL),
        new TestCaseData("1.8654637783131349E-5", 0x3EF38F90F278D788UL),
        new TestCaseData("8.5000000000000006E-2", 0x3FB5C28F5C28F5C3UL),
        new TestCaseData("7.0974999999999996E-2", 0x3FB22B6AE7D566CFUL),
        new TestCaseData("5.8427534995500012E-2", 0x3FADEA36BFFBAECFUL),
        new TestCaseData("7.1730669903414104E-2", 0x3FB25CF0F15AF73DUL),
        new TestCaseData("68.738276465294106", 0x40512F3FEBEE760FUL),
        new TestCaseData("0.64949595641478497", 0x3FE4C8ABBE75F19BUL),
        new TestCaseData("1176.4705882352939", 0x409261E1E1E1E1E1UL),
        new TestCaseData("11.576328541025971", 0x4027271488D6EA77UL),
        new TestCaseData("16.076328541025969", 0x4030138A446B753BUL),
        new TestCaseData("0.31961180303614128", 0x3FD47485105D2C29UL),
        new TestCaseData("4.8196118030361417", 0x401347485105D2C3UL),
        new TestCaseData("4.4999999999999998E-2", 0x3FA70A3D70A3D70AUL),
        new TestCaseData("3.2799597532906931E-3", 0x3F6ADE92FBEC2C61UL),
        new TestCaseData("8.6394960319995793", 0x4021476C061EB84DUL),
        new TestCaseData("-0.10346701834730029", 0xBFBA7CD084042B5EUL),
        new TestCaseData("1.3759264999163581E-6", 0x3EB7158F2F850001UL),
        new TestCaseData("-0.86136133483801225", 0xBFEB9045A56561A2UL),
        new TestCaseData("-3.5342534968048879E-2", 0xBFA2186AAFB0E635UL),
        new TestCaseData("-1.255439488689891E-5", 0xBEEA5416E50CFFFFUL),
        new TestCaseData("0.91836691248611024", 0x3FED634301DB671EUL),
        new TestCaseData("8.1633087513889813E-2", 0x3FB4E5E7F124C714UL),
        new TestCaseData("71.923671458974027", 0x4051FB1D6EE522B1UL),
        new TestCaseData("5.8332631425899999", 0x40175542EEE98B81UL),
        new TestCaseData("5.7165978797381998", 0x4016DDCBD5A76F1DUL),
        new TestCaseData("117.64705882352941", 0x405D696969696969UL),
        new TestCaseData("0.21418159555235439", 0x3FCB6A4D7226B70CUL),
        new TestCaseData("251.97834770865211", 0x406F7F4E9FDA98E0UL),
        new TestCaseData("2.3553612818440761E-5", 0x3EF8B29FF2B85BD7UL),
        new TestCaseData("1.0256220499444959", 0x3FF068F2AAA916E4UL),
        new TestCaseData("1.4214999981278845E-3", 0x3F574A340034CF9BUL),
        new TestCaseData("89648.105656933447", 0x40F5E301B0C5531CUL),
        new TestCaseData("1.3123206336756597E-2", 0x3F8AE056F047BD92UL),
    };

    [TestCaseSource(nameof(ExcelNumbers))]
    [Issue("38")]
    public void Reads_number_that_excel_saved_as_nearest_double(string text, ulong expectedBits)
    {
        AssertSheetData(
            $"""
            <sheetData>
              <row r="1">
                <c r="A1"><v>{text}</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(expectedBits, (ulong)BitConverter.DoubleToInt64Bits(ws.Cell("A1").GetDouble())));
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
    [TestCase("8.9884656743115794E+307", 0x7FE0000000000000UL)]
    [TestCase("9007199254740991.9999999999999999999", 0x4340000000000000UL)]
    [TestCase("1.1102230246251565E-16", 0x3CA0000000000000UL)]
    [TestCase("2.2250738585072014E-308", 0x0010000000000000UL)]
    [TestCase("2.2250738585072011E-308", 0x000FFFFFFFFFFFFFUL)]
    [TestCase("4.9406564584124654E-324", 0x0000000000000001UL)]
    [TestCase("2.4703282292062328E-324", 0x0000000000000001UL)]
    [TestCase("2.4703282292062327E-324", 0x0000000000000000UL)]
    [TestCase("1e-400", 0x0000000000000000UL)]
    [TestCase("1e-999999", 0x0000000000000000UL)]
    [TestCase("-0", 0x8000000000000000UL)]
    [TestCase("0.000", 0x0000000000000000UL)]
    [TestCase(".5", 0x3FE0000000000000UL)]
    [TestCase("5.", 0x4014000000000000UL)]
    [TestCase("-00012.50e-1", 0xBFF4000000000000UL)]
    [TestCase(" \t+1.5E+3 ", 0x4097700000000000UL)]
    [Issue("38")]
    public void Reads_number_text_at_or_near_midpoint_as_nearest_double(string text, ulong expectedBits)
    {
        // A text at the midpoint between two doubles, or with more digits than a double keeps, needs exact rounding.
        AssertSheetData(
            $"""
            <sheetData>
              <row r="1">
                <c r="A1"><v>{text}</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(expectedBits, (ulong)BitConverter.DoubleToInt64Bits(ws.Cell("A1").GetDouble())));
    }

    [TestCase("2.7755575615628914E-17", 0x3C80000000000000UL)]
    [TestCase("1.2345678901234567E+200", 0x6979CE4AE6F82488UL)]
    [TestCase("4.9406564584124654E-320", 0x0000000000002710UL)]
    [TestCase("1234567890123456789e-345", 0x0000000000000000UL)]
    [TestCase("9999999999999999999e-342", 0x0000000000000002UL)]
    [TestCase("1e-330", 0x0000000000000000UL)]
    [TestCase("1e-323", 0x0000000000000002UL)]
    [Issue("47")]
    public void Reads_number_text_with_large_decimal_scale_as_nearest_double(string text, ulong expectedBits)
    {
        // Excel writes such texts for the rounding errors of formulas. Some of the values are subnormal.
        AssertSheetData(
            $"""
            <sheetData>
              <row r="1">
                <c r="A1"><v>{text}</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(expectedBits, (ulong)BitConverter.DoubleToInt64Bits(ws.Cell("A1").GetDouble())));
    }

    [Test]
    [Issue("38")]
    public void Reads_long_number_text_with_all_digits()
    {
        // A non-zero digit after many zeros decides the rounding of a text at the midpoint between two doubles.
        var tie = "9007199254740993." + new string('0', 900);
        var aboveTie = tie + "1";
        var longInteger = new string('1', 850) + "e-845";
        AssertSheetData(
            $"""
            <sheetData>
              <row r="1">
                <c r="A1"><v>{tie}</v></c>
                <c r="B1"><v>{aboveTie}</v></c>
                <c r="C1"><v>{longInteger}</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(9007199254740992d, ws.Cell("A1").GetDouble());
                Assert.AreEqual(9007199254740994d, ws.Cell("B1").GetDouble());
                Assert.AreEqual(0x40C5B38E38E38E39UL, (ulong)BitConverter.DoubleToInt64Bits(ws.Cell("C1").GetDouble()));
            });
    }

    [Test]
    [Issue("38")]
    public void Reads_17_digit_text_of_random_numbers_as_same_number()
    {
        // A text with 17 significant digits identifies one double.
        var random = new Random(12345);
        var values = new double[20000];
        var bytes = new byte[8];
        for (var i = 0; i < values.Length; ++i)
        {
            if (i % 2 == 0)
            {
                values[i] = (random.NextDouble() - 0.5) * Math.Pow(10, random.Next(-10, 11));
                continue;
            }

            double value;
            do
            {
                random.NextBytes(bytes);
                value = BitConverter.ToDouble(bytes, 0);
            } while (double.IsNaN(value) || double.IsInfinity(value) || value == 0);

            values[i] = value;
        }

        var xml = new StringBuilder("<sheetData>");
        for (var i = 0; i < values.Length; ++i)
            xml.Append($"<row r=\"{i + 1}\"><c r=\"A{i + 1}\"><v>{values[i].ToString("G17", CultureInfo.InvariantCulture)}</v></c></row>");

        xml.Append("</sheetData>");

        AssertSheetData(xml.ToString(), (ws, _) =>
        {
            var differences = Enumerable.Range(0, values.Length)
                .Where(i => BitConverter.DoubleToInt64Bits(ws.Cell(i + 1, 1).GetDouble()) != BitConverter.DoubleToInt64Bits(values[i]))
                .Select(i => $"{values[i].ToString("G17", CultureInfo.InvariantCulture)} {BitConverter.DoubleToInt64Bits(values[i]):X16} {BitConverter.DoubleToInt64Bits(ws.Cell(i + 1, 1).GetDouble()):X16}")
                .ToList();
            Assert.IsEmpty(differences, $"{differences.Count} differences: {string.Join("; ", differences.Take(40))}");
        });
    }

    [TestCase("1", true)]
    [TestCase("TRUE", true)]
    [TestCase("true", true)]
    [TestCase("0", false)]
    [TestCase("FALSE", false)]
    public void Reads_boolean(string xmlValue, bool expected)
    {
        AssertSheetData(
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" t="b"><v>{xmlValue}</v></c>
               </row>
             </sheetData>
             """,
            (ws, _) => Assert.AreEqual(expected, ws.Cell("A1").GetBoolean()));
    }

    [Test]
    public void Reads_error()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="e"><v>#DIV/0!</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(XLError.DivisionByZero, ws.Cell("A1").GetError()));
    }

    [Test]
    public void Shared_string_uses_sst()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="s"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual("world", ws.Cell("A1").GetString()),
            sst: ["hello", "world"]);
    }

    [Test]
    public void Shared_string_out_of_range_is_empty()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="s"><v>9</v></c>
                <c r="B1" t="s"><v>nope</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(string.Empty, ws.Cell("A1").GetString());
                Assert.AreEqual(string.Empty, ws.Cell("B1").GetString());
            },
            sst: ["only"]);
    }

    [Test]
    public void Shared_string_without_value_is_empty_string()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="s"/>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(string.Empty, ws.Cell("A1").GetString()));
    }

    [Test]
    public void Inline_string()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="inlineStr">
                  <is>
                    <t>Hello_x0009_!</t>
                  </is>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("Hello\t!", ws.Cell("A1").GetString());
                Assert.False(ws.Cell("A1").ShareString);
            });
    }

    [Test]
    public void Inline_rich_text()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="inlineStr">
                  <is>
                    <r>
                      <t>Hi</t>
                    </r>
                  </is>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.True(ws.Cell("A1").HasRichText);
                Assert.AreEqual("Hi", ws.Cell("A1").GetRichText().Text);
                Assert.False(ws.Cell("A1").ShareString);
            });
    }

    [Test]
    public void Formula_that_results_in_string_type_has_cell_type_formula_string()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="str">
                  <f>T("x")</f>
                  <v>x</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("T(\"x\")", ws.Cell("A1").FormulaA1);
                Assert.AreEqual("x", ws.Cell("A1").GetString());
                Assert.False(ws.Cell("A1").ShareString);
            });
    }

    [Issue("ClosedXML/ClosedXML#1378")]
    [TestCase("2006-02-03T09:30:00.000", "2006-02-03T09:30:00")]
    [TestCase("2006-02-03T09:30:15", "2006-02-03T09:30:15")]
    [TestCase("2006-02-03T09:30:15.5", "2006-02-03T09:30:15.5")]
    [TestCase("2006-02-03T09:30", "2006-02-03T09:30:00")]
    [TestCase("2006-02-03", "2006-02-03T00:00:00")]
    [TestCase(" 2006-02-03 ", "2006-02-03T00:00:00")]
    public void Reads_iso_date(string xmlValue, string expectedIso)
    {
        AssertSheetData(
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" t="d"><v>{xmlValue}</v></c>
               </row>
             </sheetData>
             """,
            (ws, _) => Assert.AreEqual(DateTime.Parse(expectedIso), ws.Cell("A1").GetDateTime()));
    }

    [Issue("ClosedXML/ClosedXML#1378")]
    [TestCase("38751.5", "2006-02-03T12:00:00")]
    [TestCase("38751", "2006-02-03T00:00:00")]
    public void Reads_serial_date_in_date_cell(string xmlValue, string expectedIso)
    {
        // Some producers write a serial date number into a date cell. The value is still a date.
        AssertSheetData(
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" t="d"><v>{xmlValue}</v></c>
               </row>
             </sheetData>
             """,
            (ws, _) =>
            {
                Assert.AreEqual(XLDataType.DateTime, ws.Cell("A1").DataType);
                Assert.AreEqual(DateTime.Parse(expectedIso, CultureInfo.InvariantCulture), ws.Cell("A1").GetDateTime());
            });
    }

    [Issue("ClosedXML/ClosedXML#1378")]
    [Test]
    public void Date_cell_without_date_or_number_throws()
    {
        var ex = Assert.Throws<FormatException>(() => AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="d"><v>yesterday</v></c>
              </row>
            </sheetData>
            """,
            (_, _) => { }));
        StringAssert.Contains("yesterday", ex.Message);
    }

    [Test]
    public void Unknown_cell_type_throws()
    {
        Assert.Throws<FormatException>(() => AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="nope"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, _) => { }));
    }

    [Test]
    public void Normal_formula_with_cached_value()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <f>1+2</f>
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.AreEqual(3.0, ws.Cell("A1").CachedValue);
                Assert.False(ws.Cell("A1").ShareString);
                Assert.False(ws.Cell("A1").NeedsRecalculation);
            });
    }

    [Test]
    public void Formula_without_value_is_dirty()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <f>1+2</f>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.True(ws.Cell("A1").NeedsRecalculation);
            });
    }

    [Test]
    public void Array_formula_applies_to_range()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <f t="array" ref="A1:B1">1+2</f>
                  <v>3</v>
                </c>
                <c r="B1">
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.True(ws.Cell("A1").HasArrayFormula);
                Assert.True(ws.Cell("B1").HasArrayFormula);
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.AreEqual("1+2", ws.Cell("B1").FormulaA1);
                Assert.AreEqual("A1:B1", ws.Cell("A1").FormulaReference.ToStringRelative());
                Assert.AreEqual("A1:B1", ws.Cell("B1").FormulaReference.ToStringRelative());
            });
    }

    [Test]
    public void Array_formula_text_results_are_not_shared_strings()
    {
        // The text result of a formula is stored in the sheet, not in the shared string table.
        // This applies to each cell of the array, not only to the master cell.
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="str">
                  <f t="array" ref="A1:B2">"a"</f>
                  <v>a</v>
                </c>
                <c r="B1" t="str"><v>a</v></c>
              </row>
              <row r="2">
                <c r="A2" t="str"><v>a</v></c>
                <c r="B2" t="str"><v>a</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                foreach (var address in new[] { "A1", "B1", "A2", "B2" })
                {
                    Assert.True(ws.Cell(address).HasArrayFormula, address);
                    Assert.False(ws.Cell(address).ShareString, address);
                }
            });
    }

    [Test]
    public void Shared_formula_translates_relative_references()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="B1">
                  <f t="shared" si="0">A1+1</f>
                  <v>2</v>
                </c>
                <c r="C1">
                  <f t="shared" si="0"/>
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("A1+1", ws.Cell("B1").FormulaA1);
                Assert.AreEqual("B1+1", ws.Cell("C1").FormulaA1);
                Assert.False(ws.Cell("B1").ShareString);
                Assert.False(ws.Cell("C1").ShareString);
            });
    }

    [Test]
    public void Data_table_formula_1d_row()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="2">
                <c r="B2">
                  <f t="dataTable" ref="B2:D2" dt2D="0" dtr="1" r1="A1" del1="1"/>
                </c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var formula = ws.Cell("B2")!.Formula;
                Assert.AreEqual(FormulaType.DataTable, formula.Type);
                Assert.True(formula.IsRowDataTable);
                Assert.False(formula.Is2DDataTable);
                Assert.True(formula.Input1Deleted);
                Assert.AreEqual(new Point(1, 1), formula.Input1);
                Assert.AreEqual("B2:D2", formula.Range.ToString());
            });
    }

    [Test]
    public void Data_table_formula_2d()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="2">
                <c r="B2">
                  <f t="dataTable" ref="B2:C3" dt2D="1" r1="A1" r2="A2" del2="1"/>
                </c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var formula = ws.Cell("B2")!.Formula;
                Assert.AreEqual(FormulaType.DataTable, formula.Type);
                Assert.True(formula.Is2DDataTable);
                Assert.False(formula.Input1Deleted);
                Assert.True(formula.Input2Deleted);
                Assert.AreEqual(new Point(1, 1), formula.Input1);
                Assert.AreEqual(new Point(2, 1), formula.Input2);
            });
    }

    [Test]
    public void Row_properties()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1" ht="20" hidden="1" collapsed="1" outlineLevel="2" ph="1" customFormat="1" s="0"
                   x14ac:dyDescent="0.25">
                <c r="A1"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var row = ws.Row(1);
                Assert.AreEqual(20, row.Height);
                Assert.True(row.IsHidden);
                Assert.True(row.Collapsed);
                Assert.AreEqual(2, row.OutlineLevel);
                Assert.True(row.ShowPhonetic);
                Assert.AreEqual(0.25, row.DyDescent);
                Assert.AreSame(ws.Workbook.Styles.DefaultCellFormat, row.FormatValue);
            });
    }

    [Test]
    public void Row_without_ht_uses_sheet_row_height()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(ws.RowHeight, ws.Row(1).Height));
    }

    [Test]
    public void Cell_style_phonetic_and_meta_indexes()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" s="0" ph="1" cm="3" vm="4"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var cell = ws.Cell("A1")!;
                Assert.AreSame(ws.Workbook.Styles.DefaultCellFormat, cell.FormatValue);
                Assert.True(cell.ShowPhonetic);
                Assert.AreEqual(3u, cell.CellMetaIndex);
                Assert.AreEqual(4u, cell.ValueMetaIndex);
            });
    }

    [Test]
    public void Skips_extLst_on_row_and_cell()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <v>1</v>
                  <extLst>
                    <ext uri="{12345678-1234-1234-1234-1234567890AB}"/>
                  </extLst>
                </c>
                <c r="B1"><v>2</v></c>
                <extLst>
                  <ext uri="{12345678-1234-1234-1234-1234567890AB}"/>
                </extLst>
              </row>
              <row r="2">
                <c r="A2"><v>3</v></c>
              </row>
            </sheetData>
            """,
            (_, _) => { });
    }

    [Test]
    public void Number_with_date_format_is_datetime()
    {
        using var wb = new XLWorkbook();
        var dateTimeFormat = wb.Styles.NumberFormats[(int)XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes];
        var ws = wb.AddWorksheet();
        var dateFormat = wb.Styles.GetRegisteredCellFormat(
            wb.Styles.DefaultCellFormat,
            f => f with { NumberFormat = dateTimeFormat });
        var xfId = wb.Styles.CellFormats[dateFormat];

        LoadSheetData(
            wb,
            ws,
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" s="{xfId}" t="n"><v>44927</v></c>
               </row>
             </sheetData>
             """);

        Assert.AreEqual(XLDataType.DateTime, ws.Cell("A1").DataType);
        Assert.AreEqual(DateTime.FromOADate(44927), ws.Cell("A1").GetDateTime());
    }

    [Test]
    public void Number_with_time_format_is_timespan()
    {
        using var wb = new XLWorkbook();
        var timeSpanFormat = wb.Styles.NumberFormats[(int)XLPredefinedFormat.DateTime.Hour24MinutesSeconds];
        var ws = wb.AddWorksheet();
        var timeFormat = wb.Styles.GetRegisteredCellFormat(
            wb.Styles.DefaultCellFormat,
            f => f with { NumberFormat = timeSpanFormat });
        var xfId = wb.Styles.CellFormats[timeFormat];

        LoadSheetData(
            wb,
            ws,
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" s="{xfId}" t="n"><v>0.5</v></c>
               </row>
             </sheetData>
             """);

        Assert.AreEqual(XLDataType.TimeSpan, ws.Cell("A1").DataType);
        Assert.AreEqual(TimeSpan.FromDays(0.5), ws.Cell("A1").GetTimeSpan());
    }

    [Test]
    public void Date1904_shifts_datetime()
    {
        using var wb = new XLWorkbook();
        wb.Use1904DateSystem = true;
        var ws = wb.AddWorksheet();
        LoadSheetData(
            wb,
            ws,
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="d"><v>2006-01-01</v></c>
              </row>
            </sheetData>
            """);

        Assert.AreEqual(new DateTime(2006, 1, 1).AddDays(1462), ws.Cell("A1").GetDateTime());
    }

    [Test]
    [Issue("73")]
    public void Cell_value_decodes_escaped_characters_and_reads_long_text()
    {
        var longText = new string('x', 200);
        AssertSheetData(
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" t="str"><f>A2</f><v>a_x0041_&amp;<![CDATA[b]]></v></c>
                 <c r="B1" t="str"><f>A3</f><v>{longText}</v></c>
                 <c r="C1"><v> 1_x0035_ </v></c>
                 <c r="D1" t="str"><f>A4</f><v>short</v></c>
               </row>
             </sheetData>
             """,
            (ws, _) =>
            {
                Assert.AreEqual("aA&b", ws.Cell("A1").GetString());
                Assert.AreEqual(longText, ws.Cell("B1").GetString());
                Assert.AreEqual(15, ws.Cell("C1").GetDouble());
                Assert.AreEqual("short", ws.Cell("D1").GetString());
            });
    }

    [Test]
    [Issue("73")]
    public void Explicit_default_cell_and_formula_types_are_read()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="n">
                  <f t="normal">1+2</f>
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.AreEqual(3.0, ws.Cell("A1").CachedValue);
            });
    }

    [Test]
    [Issue("73")]
    public void Unknown_formula_type_throws()
    {
        Assert.Throws<NotSupportedException>(() => AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1"><f t="other">1</f><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, _) => { }));
    }

    [Test]
    [Issue("73")]
    public void Inline_text_with_only_phonetic_runs_is_rich_text()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="inlineStr">
                  <is>
                    <t>abc</t>
                    <rPh sb="0" eb="1"><t>A</t></rPh>
                    <phoneticPr fontId="0"/>
                  </is>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                var richText = ws.Cell("A1").GetRichText();
                Assert.AreEqual("abc", richText.Text);
                Assert.True(richText.HasPhonetics);
                Assert.AreEqual(1, richText.Phonetics.Count);
            });
    }

    private static void AssertSheetData(
        string sheetDataXml,
        Action<IXLWorksheet, XLWorksheet> assert,
        IReadOnlyList<StringItem> sst = null)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        LoadSheetData(wb, ws, sheetDataXml, sst);
        assert(ws, ws);
    }

    private static void LoadSheetData(
        XLWorkbook wb,
        IXLWorksheet ws,
        string sheetDataXml,
        IReadOnlyList<StringItem> sst = null)
    {
        using var xmlReader = OpenWorksheet(sheetDataXml);
        var parsed = new SheetDataReader(xmlReader, (XLWorksheet)ws, wb.Styles, [.. sst ?? []])
            .ParseCtSheetData("sheetData", Ns);
        Assert.True(parsed.IsSuccess);
    }

    private static XmlTreeReader OpenWorksheet(string innerXml)
    {
        var xml = $"""
                   <worksheet xmlns="{Ns}" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
                     {innerXml}
                   </worksheet>
                   """;
        var stream = new MemoryStream(XLHelper.NoBomUTF8.GetBytes(xml));
        var xmlReader = new XmlTreeReader(stream, XmlToEnumMapper.Instance, true);
        xmlReader.Open("worksheet", Ns);
        return xmlReader;
    }
}
