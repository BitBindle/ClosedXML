// Keep this file CodeMaid organised and cleaned
using System;
using System.Globalization;
using ClosedXML.Utils;

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
            // The serializer calls this method for each number cell. Format into a stack buffer,
            // so that only the text that the method returns is allocated.
            // The longest text, for example "-1.7976931348623157E+308", has 24 characters.
            Span<char> buffer = stackalloc char[32];
            if (TryFormatRoundTrip(value, buffer, "G15", out var length) ||
                TryFormatRoundTrip(value, buffer, "G16", out length))
                return new string(buffer.Slice(0, length));

            // 17 significant digits are sufficient for each double value.
            return value.ToString("G17", CultureInfo.InvariantCulture);

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
