using System;
using System.Globalization;
using System.Numerics;
using System.Text;

namespace ClosedXML.Utils
{
    /// <summary>
    /// Formats a <see cref="double"/> or a <see cref="float"/> with the fewest significant digits
    /// that parse back to the same value. The result is the same on each runtime.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The formatter tries 15, 16 and 17 significant digits for a <see cref="double"/>, and 7, 8 and 9
    /// significant digits for a <see cref="float"/>. It returns the first text that parses back to
    /// the same value. The text is the same as the text that .NET Core 3.0 and later give for the
    /// <c>G15</c>, <c>G16</c>, <c>G17</c>, <c>G7</c>, <c>G8</c> and <c>G9</c> formats with the
    /// invariant culture.
    /// </para>
    /// <para>
    /// .NET Framework (and Mono) can give a different digit at the last position of these formats,
    /// and their parsers can give a different value for a long text. Thus the formatter does not use
    /// the formatter or the parser of the runtime. It calculates the digits from the exact binary
    /// value with <see cref="BigInteger"/> arithmetic, and it rounds a half to the even digit.
    /// </para>
    /// </remarks>
    internal static class RoundTripNumberFormatter
    {
        private const int DoubleFractionBits = 52;
        private const int DoubleExponentBias = 1075; // 1023 + 52
        private const int FloatFractionBits = 23;
        private const int FloatExponentBias = 150; // 127 + 23

        /// <summary>
        /// Integers that are smaller than this limit have 15 digits or less. Thus they are exact,
        /// and the <c>G15</c> format writes them as integers.
        /// </summary>
        private const double ExactIntegerLimit = 1e15;

        private static readonly BigInteger[] PowersOfTen = CreatePowersOfTen(400);

        /// <summary>
        /// Formats a <see cref="double"/> with 15, 16 or 17 significant digits.
        /// </summary>
        internal static string Format(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
                return value.ToString(CultureInfo.InvariantCulture);

            // Most numbers in a workbook are integers. The integer text is the G15 text.
            if (Math.Abs(value) < ExactIntegerLimit && Math.Floor(value) == value && value != 0)
                return ((long)value).ToString(CultureInfo.InvariantCulture);

            var bits = BitConverter.DoubleToInt64Bits(value);
            var biasedExponent = (int)((bits >> DoubleFractionBits) & 0x7FF);
            var fraction = bits & ((1L << DoubleFractionBits) - 1);
            return Format(bits < 0, biasedExponent, fraction, DoubleFractionBits, DoubleExponentBias, 15);
        }

        /// <summary>
        /// Formats a <see cref="float"/> with 7, 8 or 9 significant digits.
        /// </summary>
        internal static string Format(float value)
        {
            if (float.IsNaN(value) || float.IsInfinity(value))
                return value.ToString(CultureInfo.InvariantCulture);

            var bits = SingleToInt32Bits(value);
            var biasedExponent = (bits >> FloatFractionBits) & 0xFF;
            var fraction = bits & ((1 << FloatFractionBits) - 1);
            return Format(bits < 0, biasedExponent, fraction, FloatFractionBits, FloatExponentBias, 7);
        }

        /// <param name="isNegative">The sign bit.</param>
        /// <param name="biasedExponent">The exponent field of the binary value.</param>
        /// <param name="fraction">The fraction field of the binary value.</param>
        /// <param name="fractionBits">The number of bits in the fraction field.</param>
        /// <param name="exponentBias">The bias of the exponent, which includes the fraction bits.</param>
        /// <param name="minPrecision">The precision to try first. The formatter also tries the
        /// next two precisions. The last precision is sufficient for each value.</param>
        private static string Format(bool isNegative, int biasedExponent, long fraction, int fractionBits, int exponentBias, int minPrecision)
        {
            if (biasedExponent == 0 && fraction == 0)
                return isNegative ? "-0" : "0";

            // The value is mantissa * 2^exponent.
            long mantissa;
            int exponent;
            if (biasedExponent == 0)
            {
                // A subnormal value.
                mantissa = fraction;
                exponent = 1 - exponentBias;
            }
            else
            {
                mantissa = fraction | (1L << fractionBits);
                exponent = biasedExponent - exponentBias;
            }

            // A text parses back to the value if its number is in the rounding interval of the value.
            // The upper half of the interval is half of the distance to the next value. The lower half
            // is half of the distance to the previous value. That distance is half the usual distance
            // for a power of two. With 4 * mantissa, all these distances are integers:
            // 2 above the value and 2 (or 1 for a power of two) below the value.
            // The parser rounds a half to the even mantissa, thus the interval includes its ends
            // only for an even mantissa.
            var scaledValue = new BigInteger(mantissa) << 2;
            var scaledExponent = exponent - 2;
            var upperGap = 2;
            var lowerGap = fraction == 0 && biasedExponent > 1 ? 1 : 2;
            var includesEnds = (mantissa & 1) == 0;

            // The value is numerator / denominator in units of 10^0.
            var numerator = scaledExponent >= 0 ? scaledValue << scaledExponent : scaledValue;
            var denominator = scaledExponent >= 0 ? BigInteger.One : BigInteger.One << -scaledExponent;

            // The decimal exponent of the first significant digit: 10^power <= value < 10^(power + 1).
            var power = (int)Math.Floor(BigInteger.Log10(numerator) - BigInteger.Log10(denominator));
            if (Compare(numerator, denominator, power + 1) >= 0)
                power++;
            else if (Compare(numerator, denominator, power) < 0)
                power--;

            for (var precision = minPrecision; ; precision++)
            {
                // The digits are value / 10^scale, rounded to an integer.
                var scale = power - precision + 1;
                var scaledNumerator = scale < 0 ? numerator * PowersOfTen[-scale] : numerator;
                var scaledDenominator = scale > 0 ? denominator * PowersOfTen[scale] : denominator;
                var digits = BigInteger.DivRem(scaledNumerator, scaledDenominator, out var remainder);
                var twiceRemainder = remainder << 1;
                var halfComparison = twiceRemainder.CompareTo(scaledDenominator);
                if (halfComparison > 0 || (halfComparison == 0 && !digits.IsEven))
                    digits += BigInteger.One;

                // The distance from the value to the text, in units of 10^scale / scaledDenominator.
                // The gaps in the same units are scaledNumerator * gap / scaledValue.
                var distance = digits * scaledDenominator - scaledNumerator;
                var gap = distance.Sign > 0 ? upperGap : lowerGap;
                var comparison = (BigInteger.Abs(distance) * scaledValue).CompareTo(scaledNumerator * gap);
                var isLastPrecision = precision == minPrecision + 2;
                if (comparison < 0 || (comparison == 0 && includesEnds) || isLastPrecision)
                {
                    var textPower = power;
                    if (digits == PowersOfTen[precision])
                    {
                        // The rounding added a digit, for example 9.99 became 10.0.
                        digits = PowersOfTen[precision - 1];
                        textPower++;
                    }

                    return FormatDigits(isNegative, digits, textPower, precision);
                }
            }
        }

        /// <summary>
        /// Compares numerator / denominator with 10^power.
        /// </summary>
        private static int Compare(BigInteger numerator, BigInteger denominator, int power)
        {
            return power >= 0
                ? numerator.CompareTo(denominator * PowersOfTen[power])
                : (numerator * PowersOfTen[-power]).CompareTo(denominator);
        }

        /// <summary>
        /// Writes the digits in the layout of the <c>G</c> format of .NET Core with the invariant
        /// culture.
        /// </summary>
        /// <param name="isNegative">Write a minus sign.</param>
        /// <param name="digits">The significant digits, as an integer.</param>
        /// <param name="power">The decimal exponent of the first digit.</param>
        /// <param name="precision">The precision of the <c>G</c> format.</param>
        private static string FormatDigits(bool isNegative, BigInteger digits, int power, int precision)
        {
            var digitText = digits.ToString(CultureInfo.InvariantCulture).TrimEnd('0');
            var text = new StringBuilder(32);
            if (isNegative)
                text.Append('-');

            if (power > -5 && power < precision)
            {
                if (power >= 0)
                {
                    var integerLength = power + 1;
                    if (digitText.Length <= integerLength)
                    {
                        text.Append(digitText).Append('0', integerLength - digitText.Length);
                    }
                    else
                    {
                        text.Append(digitText, 0, integerLength)
                            .Append('.')
                            .Append(digitText, integerLength, digitText.Length - integerLength);
                    }
                }
                else
                {
                    text.Append("0.").Append('0', -power - 1).Append(digitText);
                }
            }
            else
            {
                text.Append(digitText[0]);
                if (digitText.Length > 1)
                    text.Append('.').Append(digitText, 1, digitText.Length - 1);

                text.Append(power < 0 ? "E-" : "E+")
                    .Append(Math.Abs(power).ToString("00", CultureInfo.InvariantCulture));
            }

            return text.ToString();
        }

        private static BigInteger[] CreatePowersOfTen(int count)
        {
            var powers = new BigInteger[count];
            powers[0] = BigInteger.One;
            for (var i = 1; i < count; i++)
                powers[i] = powers[i - 1] * 10;

            return powers;
        }

        private static int SingleToInt32Bits(float value)
        {
#if NETSTANDARD2_1_OR_GREATER
            return BitConverter.SingleToInt32Bits(value);
#else
            return BitConverter.ToInt32(BitConverter.GetBytes(value), 0);
#endif
        }
    }
}
