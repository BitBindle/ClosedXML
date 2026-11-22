using ClosedXML.Excel;
using NUnit.Framework;
using System;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class XLHelperTests
    {
        [Test]
        public void IsValidColumnTest()
        {
            Assert.AreEqual(false, XLHelper.IsValidColumn(""));
            Assert.AreEqual(false, XLHelper.IsValidColumn("1"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("A1"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("AA1"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("A"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("AA"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("AAA"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("Z"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("ZZ"));
            Assert.AreEqual(true, XLHelper.IsValidColumn("XFD"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("ZAA"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("XZA"));
            Assert.AreEqual(false, XLHelper.IsValidColumn("XFZ"));
        }

        [Test]
        [SetCulture("tr-TR")]
        [Issue("ClosedXML/ClosedXML#2901")]
        public void IsValidColumn_ignores_current_culture_for_lowercase_i()
        {
            // The Turkish culture changes the upper case of "i" to "İ" (U+0130), which is not a column letter.
            Assert.That("i".ToUpper(), Is.Not.EqualTo("I"), "The test culture must change the case of 'i'.");

            Assert.That(XLHelper.IsValidColumn("i"), Is.True);
            Assert.That(XLHelper.IsValidColumn("ai"), Is.True);
            Assert.That(XLHelper.IsValidColumn("iii"), Is.True);
            Assert.That(XLHelper.IsValidA1Address("i5"), Is.True);
        }

        [Test]
        [SetCulture("tr-TR")]
        [Issue("ClosedXML/ClosedXML#2901")]
        public void Range_with_lowercase_i_column_is_valid_in_any_culture()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            var range = ws.Range("i5:j6");

            Assert.That(range.RangeAddress.ToStringRelative(), Is.EqualTo("I5:J6"));
        }

        [Test]
        public void ReplaceRelative1()
        {
            string result = XLHelper.ReplaceRelative("A1", 2, "B");
            Assert.AreEqual("B2", result);
        }

        [Test]
        public void ReplaceRelative2()
        {
            string result = XLHelper.ReplaceRelative("$A1", 2, "B");
            Assert.AreEqual("$A2", result);
        }

        [Test]
        public void ReplaceRelative3()
        {
            string result = XLHelper.ReplaceRelative("A$1", 2, "B");
            Assert.AreEqual("B$1", result);
        }

        [Test]
        public void ReplaceRelative4()
        {
            string result = XLHelper.ReplaceRelative("$A$1", 2, "B");
            Assert.AreEqual("$A$1", result);
        }

        [Test]
        public void ReplaceRelative5()
        {
            string result = XLHelper.ReplaceRelative("1:1", 2, "B");
            Assert.AreEqual("2:2", result);
        }

        [Test]
        public void ReplaceRelative6()
        {
            string result = XLHelper.ReplaceRelative("$1:1", 2, "B");
            Assert.AreEqual("$1:2", result);
        }

        [Test]
        public void ReplaceRelative7()
        {
            string result = XLHelper.ReplaceRelative("1:$1", 2, "B");
            Assert.AreEqual("2:$1", result);
        }

        [Test]
        public void ReplaceRelative8()
        {
            string result = XLHelper.ReplaceRelative("$1:$1", 2, "B");
            Assert.AreEqual("$1:$1", result);
        }

        [Test]
        public void ReplaceRelative9()
        {
            string result = XLHelper.ReplaceRelative("A:A", 2, "B");
            Assert.AreEqual("B:B", result);
        }

        [Test]
        public void ReplaceRelativeA()
        {
            string result = XLHelper.ReplaceRelative("$A:A", 2, "B");
            Assert.AreEqual("$A:B", result);
        }

        [Test]
        public void ReplaceRelativeB()
        {
            string result = XLHelper.ReplaceRelative("A:$A", 2, "B");
            Assert.AreEqual("B:$A", result);
        }

        [Test]
        public void ReplaceRelativeC()
        {
            string result = XLHelper.ReplaceRelative("$A:$A", 2, "B");
            Assert.AreEqual("$A:$A", result);
        }

        [TestCase("Sheet1", "Sheet1")]
        [TestCase("O'Brien's sales", "O'Brien's sales")]
        [TestCase(" data # ", " data # ")]
        [TestCase("data $1.00", "data $1.00")]
        [TestCase("data ", "data?")]
        [TestCase("abc def", "abc/def")]
        [TestCase("data 0 ", "data[0]")]
        [TestCase("data ", "data*")]
        [TestCase("abc def", "abc\\def")]
        [TestCase(" data", "'data")]
        [TestCase("data ", "data'")]
        [TestCase("d'at'a", "d'at'a")]
        [TestCase("sheet a4", "sheet:a4")]
        [TestCase("null", null)]
        [TestCase("empty", "")]
        [TestCase("1234567890123456789012345678901", "1234567890123456789012345678901TOOLONG")]
        public void CreateSafeSheetNames(string expected, string input)
        {
            var actual = XLHelper.CreateSafeSheetName(input);
            Assert.AreEqual(expected, actual);
        }

        [TestCase("Sheet1", ExpectedResult = "Sheet1")]
        [TestCase("O'Brien's sales", ExpectedResult = "O'Brien's sales")]
        [TestCase(" data # ", ExpectedResult = " data # ")]
        [TestCase("data $1.00", ExpectedResult = "data $1.00")]
        [TestCase("data?", ExpectedResult = "data_")]
        [TestCase("abc/def", ExpectedResult = "abc_def")]
        [TestCase("data[0]", ExpectedResult = "data_0_")]
        [TestCase("data*", ExpectedResult = "data_")]
        [TestCase("abc\\def", ExpectedResult = "abc_def")]
        [TestCase("'data", ExpectedResult = "_data")]
        [TestCase("data'", ExpectedResult = "data_")]
        [TestCase("d'at'a", ExpectedResult = "d'at'a")]
        [TestCase("sheet:a4", ExpectedResult = "sheet_a4")]
        [TestCase(null, ExpectedResult = "null")]
        [TestCase("", ExpectedResult = "empty")]
        [TestCase("1234567890123456789012345678901TOOLONG", ExpectedResult = "1234567890123456789012345678901")]
        public string CreateSafeSheetNamesWithUnderscore(string input)
        {
            return XLHelper.CreateSafeSheetName(input, replaceChar: '_');
        }

        [Test]
        public void CreateSafeSheetNamesInvalidReplacementChar()
        {
            Assert.Throws<ArgumentException>(() => XLHelper.CreateSafeSheetName("abc\\def", replaceChar: ':'));
        }

        [Test]
        [Issue("77")]
        public void Column_letter_converts_to_column_number_and_back()
        {
            for (var column = XLHelper.MinColumnNumber; column <= XLHelper.MaxColumnNumber; column++)
            {
                var letter = XLHelper.GetColumnLetterFromNumber(column);
                Assert.AreEqual(column, XLHelper.GetColumnNumberFromLetter(letter));
                Assert.AreEqual(column, XLHelper.GetColumnNumberFromLetter(letter.ToLowerInvariant()));
                Assert.AreSame(letter, XLHelper.GetColumnLetterFromNumber(column));
            }

            Assert.AreEqual("A", XLHelper.GetColumnLetterFromNumber(1));
            Assert.AreEqual("Z", XLHelper.GetColumnLetterFromNumber(26));
            Assert.AreEqual("AA", XLHelper.GetColumnLetterFromNumber(27));
            Assert.AreEqual("ZZ", XLHelper.GetColumnLetterFromNumber(702));
            Assert.AreEqual("AAA", XLHelper.GetColumnLetterFromNumber(703));
            Assert.AreEqual("XFD", XLHelper.GetColumnLetterFromNumber(16384));
        }

        [TestCase("12", 12)]
        [TestCase(" 3", 3)]
        [TestCase("xfd", 16384)]
        [Issue("77")]
        public void Column_number_from_text(string text, int expectedColumn)
        {
            Assert.AreEqual(expectedColumn, XLHelper.GetColumnNumberFromLetter(text));
        }

        [TestCase("XFE")]
        [TestCase("AAAA")]
        [TestCase("A-")]
        [TestCase("\u212A")]
        [TestCase("A\u00AD")]
        [Issue("77")]
        public void Column_number_needs_ascii_column_letters(string text)
        {
            // The Kelvin sign looks like a K, and a soft hyphen is not visible. Both are not column letters.
            Assert.Throws<ArgumentOutOfRangeException>(() => XLHelper.GetColumnNumberFromLetter(text));
        }

        [TestCase("B5", true)]
        [TestCase("$B$5", true)]
        [TestCase("b$5", true)]
        [TestCase("XFD1048576", true)]
        [TestCase("A1 ", true)]
        [TestCase("$A$1                                  ", true)]
        [TestCase("$A$0                                  ", false)]
        [TestCase("A0", false)]
        [TestCase("XFE1", false)]
        [TestCase("A1048577", false)]
        [TestCase("1A", false)]
        [TestCase("ABCD1", false)]
        [TestCase("AB", false)]
        [TestCase("$", false)]
        [Issue("77")]
        public void A1_address_validation(string address, bool expected)
        {
            Assert.AreEqual(expected, XLHelper.IsValidA1Address(address));
        }

        [TestCase("1", true)]
        [TestCase("1048576", true)]
        [TestCase("0", false)]
        [TestCase("1048577", false)]
        [TestCase("A", false)]
        [Issue("77")]
        public void Row_validation(string row, bool expected)
        {
            Assert.AreEqual(expected, XLHelper.IsValidRow(row));
        }

        [TestCase("B5", 5, 2, false, false)]
        [TestCase("$B5", 5, 2, false, true)]
        [TestCase("B$5", 5, 2, true, false)]
        [TestCase("$xfd$1048576", 1048576, 16384, true, true)]
        [Issue("77")]
        public void Create_reads_row_column_and_dollar_signs(string text, int row, int column, bool fixedRow, bool fixedColumn)
        {
            var address = XLAddress.Create(text);
            Assert.AreEqual(row, address.RowNumber);
            Assert.AreEqual(column, address.ColumnNumber);
            Assert.AreEqual(fixedRow, address.FixedRow);
            Assert.AreEqual(fixedColumn, address.FixedColumn);
        }

        [TestCase("B2:C3", "B2:C3")]
        [TestCase("Sheet!$B$2:C3", "$B$2:C3")]
        [TestCase("B2:C3:D4", "B2:C3")]
        [TestCase("D4", "D4:D4")]
        [TestCase("$2:3", "A2:XFD3")]
        [TestCase("B:$D", "B1:D1048576")]
        [Issue("77")]
        public void Constructor_reads_range_address(string text, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = (XLWorksheet)wb.AddWorksheet("Sheet");
            var rangeAddress = new XLRangeAddress(ws, text);
            Assert.AreEqual(expected, rangeAddress.FirstAddress + ":" + rangeAddress.LastAddress);
        }
    }
}
