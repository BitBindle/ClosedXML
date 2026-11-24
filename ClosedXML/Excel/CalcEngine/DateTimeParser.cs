using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class DateTimeParser
    {
        private const DateTimeStyles Style = DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowInnerWhite | DateTimeStyles.AllowTrailingWhite;

        // It's highly likely that Excel has its own database of culture specific patterns for parsing.
        // Excel has it's own parser (that accepts 1900-02-29 ^_^), never seems to parse name of a day,
        // values of hours can be up to 9999 and safely overflow...
        // Although for displaying, Excel takes a cue from region setting pattern, not so for parsing (at least
        // couldn't produce observable difference by changing setting of a culture in region dialogue).
        // .NET Core and .NET Framework also produce different patterns for GetAllDateTimePatterns.
        // This is not a perfect solution by any means, but best we can do in absence of knowledge
        // what patterns Excel uses for which cultures.
        private static readonly ConcurrentDictionary<CultureInfo, (string[] Patterns, string[] PatternsWithSeconds)> CultureSpecificPatterns = new();

        private static readonly string[] TimeOfDayPatterns = { "h:m tt", "h:m t", "h:m:s tt", "h:m:s t" };

        public static bool TryParseCultureDate(string s, CultureInfo culture, out DateTime date)
        {
            var (datePatterns, datePatternsWithSeconds) = CultureSpecificPatterns.GetOrAdd(culture, static ci =>
            {
                // Patterns that look for exactly two MM/dd that aren't part of longer sequence of MMM/ddd.
                // The MM/dd matches only two digit month/day and date recognition should be more fuzzy, it
                // should recognize month/day even without leading zero. Many cultures return only MM/dd
                // (NOT M/d) on GetAllDateTimePatterns.
                const string leadingZeroMonthPattern = @"(?<!M)MM(?!M)";
                const string leadingZeroDayPattern = @"(?<!d)dd(?!d)";
                var shortDatePatterns = ci.DateTimeFormat.GetAllDateTimePatterns('d')
                    .Concat(ci.DateTimeFormat.GetAllDateTimePatterns('D'))
                    .Where(pattern => !pattern.Contains("dddd")) // It doesn't seem that Excel parser is capable of parsing day names in any culture
                    .Select(pattern => Regex.Replace(pattern, leadingZeroMonthPattern, "M")) // Recognize months even without leading zero
                    .Select(pattern => Regex.Replace(pattern, leadingZeroDayPattern, "d")) // Recognize days even without leading zero
                    // ISO8601 should be parseable in all cultures. Be more forgiving, M,d instead MM,dd. Year first with
                    // a slash (Excel en-US accepts '2011/1/15') isn't ambiguous either, so use a literal slash.
                    .Concat(new[] { "yyyy-M-d", "yyyy'/'M'/'d" })
                    .Distinct().ToArray();

                // Not sure about this, but reasonably close. Hours pattern is probably generated (e.g. 'as-IN' culture
                // has AM designator before hours in patterns, but Excel requires it to be at the end). There most likely
                // isn't a pattern to just use. Example: for en-US, Excel type coercion can transform "aug 10, 2022 14:10",
                // but every single format from CultureInfo.DateTimeFormat requires AM/PM. and two digits for minutes (thus
                // the input couldn't match in any format => excel has likely it's own logic, independent of region setting).
                var timePatterns = new[] { "h:m tt", "H:m", "h:m" };
                var patterns = shortDatePatterns.Concat(WithTime(shortDatePatterns, timePatterns)).Distinct().ToArray();

                // Excel reads a time with seconds alone (formats 19 and 21). The parse also reads it after a date.
                var secondsPatterns = new[] { "h:m:s tt", "H:m:s" };
                var patternsWithSeconds = patterns.Concat(WithTime(shortDatePatterns, secondsPatterns)).Distinct().ToArray();
                return (patterns, patternsWithSeconds);

                static IEnumerable<string> WithTime(string[] datePatterns, string[] timePatterns) => datePatterns
                    .SelectMany(datePattern => timePatterns.Select(timePattern => FormattableString.Invariant($"{datePattern} {timePattern}")));
            });

            var patterns = MayHaveDateWithSeconds(s, culture) ? datePatternsWithSeconds : datePatterns;
            return DateTime.TryParseExact(s, patterns, culture, Style, out date);
        }

        /// <summary>
        /// A date with a time with seconds has a white space between the date and the time, and two time
        /// separators. Other texts skip the patterns with seconds, because each pattern makes a failed parse slower.
        /// </summary>
        public static bool MayHaveDateWithSeconds(string s, CultureInfo culture)
        {
            var text = s.AsSpan().Trim();
            if (text.IndexOfAny(' ', '\t') < 0)
                return false;

            // A ':' in a pattern matches a ':' or the time separator of the culture.
            var separator = culture.DateTimeFormat.TimeSeparator;
            var other = separator.Length > 0 ? separator[0] : ':';
            var count = 0;
            foreach (var c in text)
            {
                if ((c == ':' || c == other) && ++count == 2)
                    return true;
            }

            return false;
        }

        public static bool TryParseTimeOfDay(string s, CultureInfo c, out DateTime timeOfDay)
        {
            if (DateTime.TryParseExact(s, TimeOfDayPatterns, c, Style, out timeOfDay))
                return true;

            if (DateTime.TryParseExact(s, TimeOfDayPatterns, CultureInfo.InvariantCulture, Style, out timeOfDay))
                return true;

            return false;
        }
    }
}
