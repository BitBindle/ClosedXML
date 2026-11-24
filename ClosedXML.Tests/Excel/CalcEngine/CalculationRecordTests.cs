using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// The records of <see cref="IXLWorkbook.RecalculateAllFormulas(ICollection{IXLCell}, ICollection{XLCalculationRecord})"/>.
    /// A record marks a formula cell whose result can depend on the locale or can differ from Excel.
    /// </summary>
    [TestFixture]
    public class CalculationRecordTests
    {
        private const string Locale = "E1:LocaleOperation";
        private const string Engine = "E1:EngineDifference";

        [OneTimeSetUp]
        public void SetCultureInfo()
        {
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
        }

        // A text with a digit to a number or a date.
        [TestCase("\"1\"+0", Locale)]
        [TestCase("IFERROR(\"12 apples\"+0,0)", Locale)]
        [TestCase("IFERROR(\"abc\"+0,0)")]
        [TestCase("DATEVALUE(\"2020-01-05\")", Locale)]
        [TestCase("IFERROR(DATEVALUE(\"abc\"),0)")]
        [TestCase("VALUE(\"1.5\")", Locale)]
        [TestCase("VALUE(\"2011/02/23\")", Locale)]
        [TestCase("VALUE(\"January 2022\")", Locale)]
        [TestCase("IFERROR(VALUE(\"abc\"),0)")]
        [TestCase("IFERROR(VALUE(\"January\"),0)")]
        // A date text with seconds, or with the name of a month and a day. A month name without a digit is no date
        // in any locale.
        [TestCase("\"15-Jan-2011 15:30:45\"+0", Locale)]
        [TestCase("TIMEVALUE(\"1/15/2011 3:30:45 PM\")", Locale)]
        [TestCase("DATEVALUE(\"JAN 1\")", Locale)]
        [TestCase("\"DEC1\"+0", Locale)]
        [TestCase("IFERROR(DATEVALUE(\"JAN\"),0)")]
        [TestCase("COUNTIF(B1:B3,\">1\")", Locale)]
        // A number or a logical value to a text.
        [TestCase("\"a\"&1", Locale)]
        [TestCase("\"a\"&TRUE", Locale)]
        [TestCase("\"a\"&\"b\"")]
        [TestCase("TEXT(1,\"0\")", Locale)]
        [TestCase("TEXT(\"abc\",\"0\")")]
        [TestCase("FIXED(1)", Locale)]
        [TestCase("DOLLAR(1)", Locale)]
        // A text read as a logical value.
        [TestCase("IF(\"TRUE\",1,2)", Locale)]
        [TestCase("IF(TRUE,1,2)")]
        // A criterion with the name of a logical or an error value, in English or in German.
        [TestCase("COUNTIF(A1:A3,\"TRUE\")", Locale)]
        [TestCase("COUNTIF(A1:A3,\"WAHR\")", Locale)]
        [TestCase("COUNTIF(A1:A3,\"#NV\")", Locale)]
        [TestCase("COUNTIF(A1:A3,\"<>FALSE\")", Locale)]
        [TestCase("COUNTIF(A1:A3,\"#N/A\")", Locale)]
        [TestCase("COUNTIF(A1:A3,\"pear\")")]
        // A text criterion on a logical value, and an order criterion on texts.
        [TestCase("COUNTIF(C1,\"x\")", Locale)]
        [TestCase("COUNTIF(A1:A3,\">b\")", Locale)]
        // An order comparison of two texts, and an equality of two texts that both have the letter i.
        [TestCase("\"a\"<\"b\"", Locale)]
        [TestCase("\"a\"=\"b\"")]
        [TestCase("\"i\"=\"I\"", Locale)]
        [TestCase("\"kiwi\"=\"KIWI\"", Locale)]
        [TestCase("\"\u0131\"<>\"\u0130\"", Locale)]
        [TestCase("\"kiwi\"=\"pear\"")]
        [TestCase("\"a\"<>\"\u0130\"")]
        [TestCase("1<2")]
        // An approximate lookup of a text.
        [TestCase("MATCH(\"b\",A1:A3,1)", Locale)]
        [TestCase("VLOOKUP(\"b\",A1:B3,2)", Locale)]
        [TestCase("MATCH(2,B1:B3,1)")]
        // An exact lookup or a criterion, when the pattern and a text both have the letter i.
        [TestCase("MATCH(\"KIWI\",D1,0)", Locale)]
        [TestCase("COUNTIF(D1,\"k?wI\")", Locale)]
        [TestCase("MATCH(\"kiwi\",A1:A3,0)")]
        [TestCase("COUNTIF(D1,\"k*\")")]
        [TestCase("MATCH(\"plum\",A1:A3,0)")]
        // The case of a text with the letter i.
        [TestCase("UPPER(\"i\")", Locale)]
        [TestCase("LOWER(\"I\")", Locale)]
        [TestCase("PROPER(\"\u0131\")", Locale)]
        [TestCase("UPPER(\"a\")&LOWER(\"B\")&PROPER(\"c\")")]
        // A lookup text or a criterion longer than 255 characters.
        [TestCase("MATCH(REPT(\"a\",256),A1:A3,0)", Engine)]
        [TestCase("MATCH(REPT(\"a\",255),A1:A3,0)")]
        [TestCase("COUNTIF(A1:A3,REPT(\"a\",256))", Engine)]
        [TestCase("COUNTIF(A1:A3,REPT(\"a\",255))")]
        // A wildcard pattern with a ~ that is not followed by *, ? or ~.
        [TestCase("COUNTIF(A1:A3,\"a~b\")", Engine)]
        [TestCase("COUNTIF(A1:A3,\"a~\")", Engine)]
        [TestCase("VLOOKUP(\"p~ear\",A1:B3,2,FALSE)", Engine)]
        [TestCase("COUNTIF(A1:A3,\"a~*~?~~\")")]
        // Each kind once for each cell, in the order of the kinds.
        [TestCase("\"1\"+\"2\"+LEN(1)+COUNTIF(A1:A3,\"a~b\")+COUNTIF(A1:A3,REPT(\"a\",256))", Locale, Engine)]
        [TestCase("COUNTIF(A1:A3,\"a~b\")+(\"1\"+0)", Locale, Engine)]
        [TestCase("1+2")]
        public void Formula_gives_records_of_its_operations(string formula, params string[] expected)
        {
            using var wb = Workbook(formula);
            var failed = new List<IXLCell>();
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(failed, records);

            CollectionAssert.IsEmpty(failed);
            CollectionAssert.AreEqual(expected, Texts(records));
        }

        [Test]
        public void Records_dont_change_values()
        {
            using var withRecords = Workbook("COUNTIF(A1:A3,\">b\")&TEXT(1.5,\"0.0\")&\"1\"*2");
            using var without = Workbook("COUNTIF(A1:A3,\">b\")&TEXT(1.5,\"0.0\")&\"1\"*2");

            withRecords.RecalculateAllFormulas(new List<IXLCell>(), new List<XLCalculationRecord>());
            without.RecalculateAllFormulas();

            Assert.AreEqual("21.52", withRecords.Worksheet(1).Cell("E1").CachedValue);
            Assert.AreEqual(without.Worksheet(1).Cell("E1").CachedValue, withRecords.Worksheet(1).Cell("E1").CachedValue);
        }

        [Test]
        public void Records_are_off_by_default()
        {
            using var wb = Workbook("\"1\"+0");
            var records = new List<XLCalculationRecord>();
            wb.RecalculateAllFormulas(new List<IXLCell>(), records);
            Assert.AreEqual(1, records.Count);

            // The recorder of the thread ends with the recalculation that asked for it.
            Assert.IsNull(ClosedXML.Excel.CalcEngine.CalculationRecorder.Current);
            wb.RecalculateAllFormulas(new List<IXLCell>());
            wb.RecalculateAllFormulas();
            wb.Worksheet(1).Cell("E1").InvalidateFormula();
            wb.RecalculateDirtyFormulas(new List<IXLCell>());
            wb.Worksheet(1).Cell("E1").InvalidateFormula();
            wb.RecalculateDirtyFormulas();
            wb.Worksheet(1).Cell("E2").FormulaA1 = "\"2\"+0";
            Assert.AreEqual(2, wb.Worksheet(1).Cell("E2").Value);
            Assert.AreEqual(1, records.Count);
        }

        [Test]
        public void Failed_cell_keeps_records_of_operations_before_failure()
        {
            using var wb = Workbook("TEXT(1,\"0\")&A2#");
            wb.Worksheet(1).Cell("F1").FormulaA1 = "\"1\"+0";
            var failed = new List<IXLCell>();
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(failed, records);

            CollectionAssert.AreEqual(new[] { "E1" }, failed.Select(c => c.Address.ToString()));
            CollectionAssert.AreEqual(new[] { Locale, "F1:LocaleOperation" }, Texts(records));
        }

        [Test]
        public void Cell_that_waits_for_a_dirty_cell_gets_one_record()
        {
            // E1 is first in the chain. Its evaluation stops at the dirty F1, and E1 is evaluated again after F1.
            using var wb = Workbook("VALUE(\"1\")+F1");
            wb.Worksheet(1).Cell("F1").FormulaA1 = "\"2\"+0";
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(new List<IXLCell>(), records);

            Assert.AreEqual(3, wb.Worksheet(1).Cell("E1").CachedValue);
            CollectionAssert.AreEquivalent(new[] { Locale, "F1:LocaleOperation" }, Texts(records));
        }

        [Test]
        public void Array_formula_gives_one_record_for_its_first_cell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("E1:E3").FormulaArrayA1 = "\"1\"+{0;1;2}";
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(new List<IXLCell>(), records);

            Assert.AreEqual(3, ws.Cell("E3").CachedValue);
            CollectionAssert.AreEqual(new[] { Locale }, Texts(records));
        }

        [Test]
        public void Record_has_the_formula_cell()
        {
            using var wb = Workbook("\"1\"+0");
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(new List<IXLCell>(), records);

            var (cell, kind) = records.Single();
            Assert.AreSame(wb.Worksheet(1), cell.Worksheet);
            Assert.AreEqual("E1", cell.Address.ToString());
            Assert.AreEqual(XLCalculationRecordKind.LocaleOperation, kind);
        }

        [Test]
        public void Collection_that_reads_cell_values_gets_the_same_records()
        {
            // E1 fails. E2 and E4 wait for E3. The collection reads the value of each recorded cell.
            using var wb = Workbook("TEXT(1,\"0\")&A2#");
            var ws = wb.Worksheet(1);
            ws.Cell("E2").FormulaA1 = "E3+1";
            ws.Cell("E3").FormulaA1 = "\"5\"+0";
            ws.Cell("E4").FormulaA1 = "E2*2";
            var failed = new List<IXLCell>();
            var records = new CallbackCollection<XLCalculationRecord>(r => _ = r.Cell.Value);

            wb.RecalculateAllFormulas(failed, records);

            CollectionAssert.AreEqual(new[] { "E1" }, failed.Select(c => c.Address.ToString()));
            CollectionAssert.AreEqual(new[] { Locale, "E3:LocaleOperation" }, Texts(records));
            Assert.AreEqual(12, ws.Cell("E4").CachedValue);
        }

        [Test]
        public void Exception_of_the_collection_stops_the_recalculation()
        {
            // An exception of the collection is not a failure of the formula, so it goes to the caller.
            using var wb = Workbook("\"1\"+COUNTIF(A1:A3,\"a~b\")");
            var added = 0;
            var records = new CallbackCollection<XLCalculationRecord>(_ =>
            {
                if (added++ == 1)
                    throw new InvalidOperationException("Add failed.");
            });
            var failed = new List<IXLCell>();

            Assert.That(() => wb.RecalculateAllFormulas(failed, records), Throws.InvalidOperationException.With.Message.EqualTo("Add failed."));

            CollectionAssert.IsEmpty(failed);
            CollectionAssert.AreEqual(new[] { Locale }, Texts(records));
            Assert.IsNull(ClosedXML.Excel.CalcEngine.CalculationRecorder.Current);
        }

        [Test]
        public void Recalculation_of_another_workbook_during_a_recalculation_gives_no_records()
        {
            using var other = Workbook("COUNTIF(A1:A3,\"a~b\")&TEXT(1,\"0\")");
            using var wb = Workbook("\"1\"+0");
            wb.Worksheet(1).Cell("E2").FormulaA1 = "\"2\"+0";
            var records = new CallbackCollection<XLCalculationRecord>(_ => other.RecalculateAllFormulas());

            wb.RecalculateAllFormulas(new List<IXLCell>(), records);

            CollectionAssert.AreEqual(new[] { Locale, "E2:LocaleOperation" }, Texts(records));
            Assert.IsTrue(records.All(r => ReferenceEquals(r.Cell.Worksheet.Workbook, wb)));
        }

        [Test]
        public void Recalculation_of_dirty_formulas_gives_records_of_calculated_cells()
        {
            using var wb = Workbook("\"1\"+B1");
            var ws = wb.Worksheet(1);
            ws.Cell("E2").FormulaA1 = "\"2\"+0";
            ws.Cell("E3").FormulaA1 = "TEXT(B1,\"0\")&A2#";
            var records = new List<XLCalculationRecord>();
            wb.RecalculateAllFormulas(new List<IXLCell>(), records);
            CollectionAssert.AreEqual(new[] { Locale, "E2:LocaleOperation", "E3:LocaleOperation" }, Texts(records));

            // E2 doesn't read B1, so it isn't dirty and gets no record. E3 fails again and keeps its record.
            ws.Cell("B1").Value = 5;
            var failed = new List<IXLCell>();
            records.Clear();
            wb.RecalculateDirtyFormulas(failed, records);

            CollectionAssert.AreEqual(new[] { "E3" }, failed.Select(c => c.Address.ToString()));
            CollectionAssert.AreEqual(new[] { Locale, "E3:LocaleOperation" }, Texts(records));
            Assert.AreEqual(6, ws.Cell("E1").CachedValue);
        }

        [Test]
        public void Clock_that_evaluates_another_workbook_gives_no_records()
        {
            using var other = new XLWorkbook();
            var clock = new CallbackClock(() => other.Evaluate("MATCH(\"a~b\",{\"x\"},0)"));
            using var wb = new XLWorkbook(new LoadOptions { TimeProvider = clock });
            wb.AddWorksheet().Cell("A1").FormulaA1 = "TODAY()+DATEVALUE(\"Mar-35\")*0";
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(new List<IXLCell>(), records);

            Assert.Greater(clock.Calls, 0);
            CollectionAssert.AreEqual(new[] { "A1:LocaleOperation" }, Texts(records));
        }

        [Test]
        public void Name_filter_that_evaluates_another_workbook_gives_no_records()
        {
            using var other = new XLWorkbook();
            var calls = 0;
            using var wb = new XLWorkbook(new LoadOptions
            {
                CalculationNameFilter = (_, _) =>
                {
                    calls++;
                    other.Evaluate("MATCH(\"a~b\",{\"x\"},0)");
                    return true;
                }
            });
            var ws = wb.AddWorksheet();
            wb.DefinedNames.Add("Rate", "2");
            ws.Cell("A1").FormulaA1 = "Rate*2";
            var records = new List<XLCalculationRecord>();

            wb.RecalculateAllFormulas(new List<IXLCell>(), records);

            Assert.Greater(calls, 0);
            Assert.AreEqual(4, ws.Cell("A1").CachedValue);
            CollectionAssert.IsEmpty(records);
        }

        [Test]
        public void Null_collection_throws()
        {
            using var wb = Workbook("1");
            Assert.Throws<ArgumentNullException>(() => wb.RecalculateAllFormulas(null!, new List<XLCalculationRecord>()));
            Assert.Throws<ArgumentNullException>(() => wb.RecalculateAllFormulas(new List<IXLCell>(), null!));
            Assert.Throws<ArgumentNullException>(() => wb.RecalculateDirtyFormulas(null!, new List<XLCalculationRecord>()));
            Assert.Throws<ArgumentNullException>(() => wb.RecalculateDirtyFormulas(new List<IXLCell>(), null!));
        }

        private static XLWorkbook Workbook(string formula)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "apple";
            ws.Cell("A2").Value = "pear";
            ws.Cell("A3").Value = "plum";
            ws.Cell("B1").Value = 1;
            ws.Cell("B2").Value = 2;
            ws.Cell("B3").Value = 3;
            ws.Cell("C1").Value = true;
            ws.Cell("D1").Value = "kiwi";
            ws.Cell("E1").FormulaA1 = formula;
            return wb;
        }

        private sealed class CallbackClock : TimeProvider
        {
            private readonly Action _onCall;

            public CallbackClock(Action onCall) => _onCall = onCall;

            public int Calls { get; private set; }

            public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;

            public override DateTimeOffset GetUtcNow()
            {
                Calls++;
                _onCall();
                return new DateTimeOffset(2030, 6, 1, 12, 0, 0, TimeSpan.Zero);
            }
        }

        private sealed class CallbackCollection<T> : System.Collections.ObjectModel.Collection<T>
        {
            private readonly Action<T> _onAdd;

            public CallbackCollection(Action<T> onAdd) => _onAdd = onAdd;

            protected override void InsertItem(int index, T item)
            {
                _onAdd(item);
                base.InsertItem(index, item);
            }
        }

        private static string[] Texts(IEnumerable<XLCalculationRecord> records) =>
            records.Select(r => $"{r.Cell.Address}:{r.Kind}").ToArray();
    }
}
