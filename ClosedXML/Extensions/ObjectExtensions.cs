// Keep this file CodeMaid organised and cleaned
using System;
using System.Globalization;

namespace ClosedXML.Excel
{
    internal static class ObjectExtensions
    {
        public static T CastTo<T>(this Object o)
        {
            return (T)Convert.ChangeType(o, typeof(T));
        }

        public static string ToInvariantString<T>(this T value) where T : struct
        {
            return value switch
            {
                sbyte v => v.ToString(CultureInfo.InvariantCulture),
                byte v => v.ToString(CultureInfo.InvariantCulture),
                short v => v.ToString(CultureInfo.InvariantCulture),
                ushort v => v.ToString(CultureInfo.InvariantCulture),
                int v => v.ToString(CultureInfo.InvariantCulture),
                uint v => v.ToString(CultureInfo.InvariantCulture),
                long v => v.ToString(CultureInfo.InvariantCulture),
                ulong v => v.ToString(CultureInfo.InvariantCulture),
                float v => ToRoundTripString(v),
                double v => ToRoundTripString(v),
                decimal v => v.ToString(CultureInfo.InvariantCulture),
                TimeSpan ts => ts.ToString("c", CultureInfo.InvariantCulture),
                DateTime d => d.ToString(CultureInfo.InvariantCulture),
                bool b => b.ToString().ToLowerInvariant(),
                _ => value.ToString()!,
            };
        }

        /// <summary>
        /// Formats a number with the fewest significant digits that parse back to the same value.
        /// The text has 15 significant digits if they are sufficient, which is the precision that
        /// Excel shows. Otherwise the text has 16 or 17 significant digits.
        /// </summary>
        internal static string ToRoundTripString(double value)
        {
            // Format into a stack buffer. The method then allocates only the text that it returns.
            Span<char> buffer = stackalloc char[RoundTripLength];
            var length = FormatRoundTrip(value, buffer);
            return new string(buffer.Slice(0, length));
        }

        /// <summary>
        /// The minimum length of a buffer for <see cref="FormatRoundTrip"/>. The longest text, for
        /// example <c>-1.7976931348623157E+308</c>, has 24 characters.
        /// </summary>
        internal const int RoundTripLength = 32;

        /// <summary>
        /// Formats a number into a buffer with the same text as <see cref="ToRoundTripString(double)"/>.
        /// The serializer calls this method for each number cell, so the method allocates no memory.
        /// </summary>
        /// <param name="value">The number to format.</param>
        /// <param name="destination">The buffer. It has at least <see cref="RoundTripLength"/> characters.</param>
        /// <returns>The number of characters in the buffer.</returns>
        internal static int FormatRoundTrip(double value, Span<char> destination)
        {
            if (TryFormatRoundTrip(value, destination, "G15", out var length) ||
                TryFormatRoundTrip(value, destination, "G16", out length))
                return length;

            // 17 significant digits are sufficient for each double value.
            if (!value.TryFormat(destination, out length, "G17", CultureInfo.InvariantCulture))
                throw new ArgumentException("The buffer is too short.", nameof(destination));

            return length;

            static bool TryFormatRoundTrip(double value, Span<char> buffer, string format, out int length)
            {
                return value.TryFormat(buffer, out length, format, CultureInfo.InvariantCulture) &&
                       double.TryParse(buffer.Slice(0, length), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) &&
                       parsed.Equals(value);
            }
        }

        /// <summary>
        /// Formats a number with the fewest significant digits that parse back to the same value.
        /// The text has 7 or 8 significant digits if they are sufficient. Otherwise the text has
        /// 9 significant digits.
        /// </summary>
        internal static string ToRoundTripString(float value)
        {
            var text = value.ToString("G7", CultureInfo.InvariantCulture);
            if (IsRoundTrip(text, value))
                return text;

            text = value.ToString("G8", CultureInfo.InvariantCulture);
            if (IsRoundTrip(text, value))
                return text;

            // 9 significant digits are sufficient for each float value.
            return value.ToString("G9", CultureInfo.InvariantCulture);

            static bool IsRoundTrip(string text, float value)
            {
                return float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) &&
                       parsed.Equals(value);
            }
        }
    }
}
