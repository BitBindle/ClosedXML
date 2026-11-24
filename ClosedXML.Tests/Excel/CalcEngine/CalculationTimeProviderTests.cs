using System;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine
{
    /// <summary>
    /// A server that calculates for a user in another time zone, or a test, needs a fixed time for the calculation.
    /// <see cref="LoadOptions.TimeProvider"/> gives that time to <c>TODAY</c>, <c>NOW</c> and to the conversion of
    /// a date text without a year or with a two-digit year.
    /// </summary>
    [TestFixture]
    [SetCulture("en-US")]
    public class CalculationTimeProviderTests
    {
        // 2030-12-31 20:30 UTC is 2031-01-01 06:30 in a time zone of UTC+10.
        private static readonly DateTimeOffset UtcNow = new(2030, 12, 31, 20, 30, 0, TimeSpan.Zero);

        private static readonly TimeZoneInfo UtcPlus10 =
            TimeZoneInfo.CreateCustomTimeZone("UTC+10", TimeSpan.FromHours(10), "UTC+10", "UTC+10");

        [Test]
        public void Today_uses_the_local_date_of_the_time_provider()
        {
            using var wb = CreateWorkbook(UtcNow, UtcPlus10);

            var actual = wb.Evaluate("TODAY()");

            Assert.AreEqual(new DateTime(2031, 1, 1).ToOADate(), actual.GetNumber());
        }

        [Test]
        public void Now_uses_the_local_time_of_the_time_provider()
        {
            using var wb = CreateWorkbook(UtcNow, UtcPlus10);

            var actual = wb.Evaluate("NOW()");

            Assert.AreEqual(new DateTime(2031, 1, 1, 6, 30, 0).ToOADate(), actual.GetNumber(), 1e-9);
        }

        [Test]
        public void Cell_formula_uses_the_time_provider_in_a_recalculation()
        {
            using var wb = CreateWorkbook(UtcNow, UtcPlus10);
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = "YEAR(TODAY())";

            wb.RecalculateAllFormulas();

            Assert.AreEqual(2031, ws.Cell("A1").CachedValue.GetNumber());
        }

        [Test]
        public void Recalculation_on_load_uses_the_time_provider()
        {
            using var stream = new MemoryStream();
            using (var source = new XLWorkbook())
            {
                source.AddWorksheet().Cell("A1").FormulaA1 = "TODAY()";
                source.SaveAs(stream);
            }

            var options = new LoadOptions
            {
                TimeProvider = new FixedTimeProvider(UtcNow, UtcPlus10),
                RecalculateAllFormulas = true,
            };
            using var wb = new XLWorkbook(stream, options);

            Assert.AreEqual(new DateTime(2031, 1, 1).ToOADate(), wb.Worksheet(1).Cell("A1").CachedValue.GetNumber());
        }

        [TestCase("\"29-Feb\"*1")]
        [TestCase("DATEVALUE(\"29-Feb\")")]
        [TestCase("\"29-February\"+0")]
        [TestCase("VALUE(\"29-February\")")]
        public void Day_and_month_text_uses_the_year_of_the_time_provider(string formula)
        {
            // Format 16 'd-mmm' has no year. 2028 is a leap year, so the date exists only in the year of the time provider.
            using var wb = CreateWorkbook(new DateTimeOffset(2028, 6, 1, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);

            var actual = wb.Evaluate(formula);

            Assert.AreEqual(new DateTime(2028, 2, 29).ToOADate(), actual.GetNumber());
        }

        [TestCase("\"JAN 1\"*1", 1, 1)]
        [TestCase("VALUE(\"DEC1\")", 12, 1)]
        [TestCase("DATEVALUE(\"Feb 29\")", 2, 29)]
        public void Month_and_day_text_uses_the_year_of_the_time_provider(string formula, int month, int day)
        {
            // The name of a month and a day has no year. 2028 is a leap year, so February 29 exists only in it.
            using var wb = CreateWorkbook(new DateTimeOffset(2028, 6, 1, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);

            var actual = wb.Evaluate(formula);

            Assert.AreEqual(new DateTime(2028, month, day).ToOADate(), actual.GetNumber());
        }

        [Test]
        public void Value_of_day_and_month_text_uses_the_year_of_the_time_provider()
        {
            using var wb = CreateWorkbook(new DateTimeOffset(2028, 6, 1, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);

            var actual = wb.Evaluate("VALUE(\"5-July\")");

            Assert.AreEqual(new DateTime(2028, 7, 5).ToOADate(), actual.GetNumber());
        }

        [Test]
        [SetCulture("th-TH")]
        public void Day_and_month_text_uses_the_year_of_the_time_provider_in_the_calendar_of_the_culture()
        {
            // The culture th-TH reads a year in the Thai Buddhist calendar (2028 is 2571).
            using var wb = CreateWorkbook(new DateTimeOffset(2028, 6, 1, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);
            var february = CultureInfo.CurrentCulture.DateTimeFormat.AbbreviatedMonthNames[1];

            var actual = wb.Evaluate($"DATEVALUE(\"29-{february}\")");

            Assert.AreEqual(new DateTime(2028, 2, 29).ToOADate(), actual.GetNumber());
        }

        [Test]
        public void Day_and_month_criteria_uses_the_year_of_the_time_provider()
        {
            using var wb = CreateWorkbook(new DateTimeOffset(2028, 6, 1, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = new DateTime(2028, 7, 5);
            ws.Cell("A2").Value = new DateTime(2027, 7, 5);

            var actual = ws.Evaluate("COUNTIF(A1:A2,\"5-Jul\")");

            Assert.AreEqual(1, actual.GetNumber());
        }

        [TestCase(2035, "\"Mar-35\"*1", 2035)]
        [TestCase(2026, "\"Mar-35\"*1", 1935)]
        [TestCase(2035, "DATEVALUE(\"March-36\")", 1936)]
        public void Month_and_year_text_uses_the_century_of_the_time_provider(int clockYear, string formula, int expectedYear)
        {
            // Format 17 'mmm-yy' keeps a two-digit year from 2030 only in the year of the time provider.
            using var wb = CreateWorkbook(new DateTimeOffset(clockYear, 6, 1, 12, 0, 0, TimeSpan.Zero), TimeZoneInfo.Utc);

            var actual = wb.Evaluate(formula);

            Assert.AreEqual(new DateTime(expectedYear, 3, 1).ToOADate(), actual.GetNumber());
        }

        [Test]
        public void Default_time_provider_is_the_system_clock()
        {
            Assert.IsNull(new LoadOptions().TimeProvider);

            using var wb = new XLWorkbook();
            var before = DateTime.Today;
            var actual = wb.Evaluate("TODAY()").GetNumber();
            var after = DateTime.Today;

            Assert.That(actual, Is.EqualTo(before.ToOADate()).Or.EqualTo(after.ToOADate()));
        }

        private static XLWorkbook CreateWorkbook(DateTimeOffset utcNow, TimeZoneInfo localTimeZone)
        {
            return new XLWorkbook(new LoadOptions { TimeProvider = new FixedTimeProvider(utcNow, localTimeZone) });
        }

        private sealed class FixedTimeProvider : TimeProvider
        {
            private readonly DateTimeOffset _utcNow;
            private readonly TimeZoneInfo _localTimeZone;

            public FixedTimeProvider(DateTimeOffset utcNow, TimeZoneInfo localTimeZone)
            {
                _utcNow = utcNow;
                _localTimeZone = localTimeZone;
            }

            public override DateTimeOffset GetUtcNow() => _utcNow;

            public override TimeZoneInfo LocalTimeZone => _localTimeZone;
        }
    }
}
