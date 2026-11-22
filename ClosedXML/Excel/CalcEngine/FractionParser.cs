using System;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// Parse a fraction for text-to-number type coercion.
    /// </summary>
    /// <remarks>
    /// The accepted text is <c>^ *([+-]?) *([0-9]+) ([0-9]{1,5})/([0-9]{1,5}) *$</c>. As in a regular
    /// expression, the <c>$</c> also accepts one line feed at the end.
    /// </remarks>
    internal static class FractionParser
    {
        private const int MaxFractionDigits = 5;

        public static bool TryParse(ReadOnlySpan<char> s, out double result)
        {
            result = default;
            var i = SkipSpaces(s, 0);

            var hasNegativeSign = false;
            if (i < s.Length && s[i] is '+' or '-')
            {
                hasNegativeSign = s[i] == '-';
                i = SkipSpaces(s, i + 1);
            }

            var wholeNumberDigits = ReadDigits(s, i);
            if (wholeNumberDigits.IsEmpty)
                return false;

            i += wholeNumberDigits.Length;
            if (i >= s.Length || s[i] != ' ')
                return false;

            var numeratorDigits = ReadDigits(s, i + 1);
            if (numeratorDigits.IsEmpty || numeratorDigits.Length > MaxFractionDigits)
                return false;

            i += 1 + numeratorDigits.Length;
            if (i >= s.Length || s[i] != '/')
                return false;

            var denominatorDigits = ReadDigits(s, i + 1);
            if (denominatorDigits.IsEmpty || denominatorDigits.Length > MaxFractionDigits)
                return false;

            i = SkipSpaces(s, i + 1 + denominatorDigits.Length);
            var isEnd = i == s.Length || (i == s.Length - 1 && s[i] == '\n');
            if (!isEnd)
                return false;

            var denominator = ParseInt(denominatorDigits);
            if (denominator == 0 || denominator > short.MaxValue)
                return false;

            var numerator = ParseInt(numeratorDigits);
            if (numerator > short.MaxValue)
                return false;

            // The whole number can have many digits, so parse it as a double.
            var wholeNumber = double.Parse(wholeNumberDigits, NumberStyles.None, CultureInfo.InvariantCulture);
            var fraction = wholeNumber + numerator / (double)denominator;
            if (!double.IsFinite(fraction))
                return false;

            result = hasNegativeSign ? -fraction : fraction;
            return true;
        }

        private static int SkipSpaces(ReadOnlySpan<char> s, int index)
        {
            while (index < s.Length && s[index] == ' ')
                index++;

            return index;
        }

        /// <summary>
        /// Get the ASCII digits at the <paramref name="index"/>. The result is empty if no digit is at the index.
        /// </summary>
        private static ReadOnlySpan<char> ReadDigits(ReadOnlySpan<char> s, int index)
        {
            var end = index;
            while (end < s.Length && char.IsAsciiDigit(s[end]))
                end++;

            return s[index..end];
        }

        /// <summary>
        /// Parse at most <see cref="MaxFractionDigits"/> ASCII digits.
        /// </summary>
        private static int ParseInt(ReadOnlySpan<char> digits)
        {
            var value = 0;
            foreach (var digit in digits)
                value = value * 10 + (digit - '0');

            return value;
        }
    }
}
