using System;
using System.Globalization;
using System.Numerics;
using System.Text;
using ClosedXML.Excel;

namespace ClosedXML.Utils;

/// <summary>
/// Converts between a <see cref="double"/> and its decimal text with integer arithmetic only.
/// The result is the IEEE 754 result, and it does not depend on the runtime.
/// </summary>
/// <remarks>
/// The <see cref="double"/> parser and formatter of .NET Framework are not exact for texts
/// with more than 15 significant digits. .NET Core 3.0 and later are exact. The netstandard2.0
/// build uses this class, so that it reads and writes the same values on each runtime.
/// </remarks>
internal static class ExactDoubleConverter
{
    private const int MantissaBits = 52;
    private const int ExponentBias = 1075;
    private const int MinBinaryExponent = -1074;
    private const int MaxBiasedExponent = 0x7FF;

    /// <summary>
    /// 17 significant digits are sufficient to identify each <see cref="double"/>.
    /// </summary>
    private const int MaxPrecision = 17;

    /// <summary>
    /// Each integer up to 2^53 is an exact <see cref="double"/>.
    /// </summary>
    private const ulong MaxExactInteger = 1UL << (MantissaBits + 1);

    /// <summary>
    /// A decimal text never needs more significant digits than this to select the nearest
    /// <see cref="double"/>. The exact value of a <see cref="double"/> has at most 767 significant digits.
    /// </summary>
    private const int MaxSignificantDigits = 800;

    /// <summary>
    /// The powers of ten that a <see cref="double"/> keeps exactly.
    /// </summary>
    private static readonly double[] ExactPowersOfTen =
    {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11,
        1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
    };

    /// <summary>
    /// The powers of ten that the conversions of most values use. The cache removes the repeated calculation.
    /// </summary>
    private static readonly BigInteger[] PowersOfTen = CreatePowersOfTen(350);

    /// <summary>
    /// Formats a number with the fewest significant digits that parse back to the same value.
    /// The text has 15 significant digits if they are sufficient, which is the precision that
    /// Excel shows. Otherwise the text has 16 or 17 significant digits. The text is the same
    /// as the <c>G15</c>, <c>G16</c> or <c>G17</c> format of .NET Core 3.0 and later.
    /// </summary>
    internal static string Format(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
            return value.ToString(CultureInfo.InvariantCulture);

        var isNegative = BitConverter.DoubleToInt64Bits(value) < 0;
        if (value == 0)
            return isNegative ? "-0" : "0";

        var absValue = Math.Abs(value);

        // An integer below 10^15 has at most 15 digits, so the G15 text is the integer itself.
        if (absValue < 1e15 && Math.Floor(absValue) == absValue)
            return ((long)value).ToString(CultureInfo.InvariantCulture);

        GetRatio(absValue, out var numerator, out var denominator);
        var valueExponent = GetDecimalExponent(absValue, numerator, denominator);

        // One division gives the first 17 digits. The rounding to 15 and 16 digits uses them and the remainder.
        var scale = valueExponent - (MaxPrecision - 1);
        if (scale >= 0)
            denominator *= PowerOfTen(scale);
        else
            numerator *= PowerOfTen(-scale);

        var truncatedDigits = (ulong)BigInteger.DivRem(numerator, denominator, out var remainder);
        var remainderComparison = (remainder << 1).CompareTo(denominator);
        for (var precision = 15; ; precision++)
        {
            var exponent = valueExponent;
            var digits = RoundToDigits(truncatedDigits, remainder.IsZero, remainderComparison, precision, ref exponent);

            // 17 significant digits are sufficient for each double value.
            if (precision == MaxPrecision || IsRoundTrip(digits, exponent - precision + 1, absValue))
                return FormatDigits(isNegative, digits, exponent, precision);
        }
    }

    /// <summary>
    /// Parses a decimal text with an optional sign, decimal point and exponent. White space
    /// can be before and after the number. The text uses the invariant culture.
    /// </summary>
    /// <remarks>
    /// If the text has a different format, or the value is too large for a <see cref="double"/>,
    /// the method uses the parser of the runtime. Then the result is the result of the runtime.
    /// </remarks>
    internal static bool TryParse(string text, out double value)
    {
        if (TryParseExact(text, out value))
            return true;

        return double.TryParse(text, XLHelper.NumberStyle, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryParseExact(string text, out double value)
    {
        value = 0;
        var position = 0;
        var end = text.Length;
        while (position < end && IsWhite(text[position]))
            position++;
        while (end > position && IsWhite(text[end - 1]))
            end--;

        var isNegative = false;
        if (position < end && (text[position] == '+' || text[position] == '-'))
        {
            isNegative = text[position] == '-';
            position++;
        }

        // The digits keep the significant digits without the leading zeros.
        var digits = new StringBuilder();
        var hasDigit = false;
        var hasDecimalPoint = false;
        var hasDiscardedNonZeroDigit = false;
        var decimalExponent = 0;
        for (; position < end; position++)
        {
            var c = text[position];
            if (c == '.')
            {
                if (hasDecimalPoint)
                    return false;

                hasDecimalPoint = true;
                continue;
            }

            if (c < '0' || c > '9')
                break;

            hasDigit = true;
            if (digits.Length == 0 && c == '0')
            {
                if (hasDecimalPoint)
                    decimalExponent--;

                continue;
            }

            if (digits.Length < MaxSignificantDigits)
            {
                digits.Append(c);
                if (hasDecimalPoint)
                    decimalExponent--;
            }
            else
            {
                hasDiscardedNonZeroDigit |= c != '0';
                if (!hasDecimalPoint)
                    decimalExponent++;
            }
        }

        if (!hasDigit)
            return false;

        if (position < end && (text[position] == 'e' || text[position] == 'E'))
        {
            position++;
            if (!TryParseExponent(text, ref position, end, out var exponent))
                return false;

            decimalExponent += exponent;
        }

        if (position != end)
            return false;

        if (!hasDiscardedNonZeroDigit)
        {
            // Without the trailing zeros, more texts have at most 15 digits for the fast calculation.
            var trailingZeros = 0;
            while (trailingZeros < digits.Length && digits[digits.Length - 1 - trailingZeros] == '0')
                trailingZeros++;

            digits.Length -= trailingZeros;
            decimalExponent += trailingZeros;
        }
        else
        {
            // A digit below all kept digits keeps the information that the value is above the kept digits.
            digits.Append('1');
            decimalExponent--;
        }

        if (digits.Length == 0 || digits.Length + decimalExponent < -330)
        {
            value = isNegative ? -0.0 : 0.0;
            return true;
        }

        // The value is 10^309 or more. The runtime decides the result of an overflow.
        if (digits.Length + decimalExponent > 309)
            return false;

        // Excel writes 17 significant digits, so most texts fit in an unsigned long.
        var smallInteger = digits.Length <= 19 ? ParseSmallInteger(digits) : ulong.MaxValue;
        double absValue;
        if (smallInteger <= MaxExactInteger && Math.Abs(decimalExponent) < ExactPowersOfTen.Length)
        {
            // The integer and the power of ten are exact doubles, and IEEE 754 rounds the
            // product or the quotient one time. Thus the result is the nearest double.
            absValue = decimalExponent >= 0
                ? smallInteger * ExactPowersOfTen[decimalExponent]
                : smallInteger / ExactPowersOfTen[-decimalExponent];
        }
        else
        {
            var integer = digits.Length <= 19
                ? new BigInteger(smallInteger)
                : BigInteger.Parse(digits.ToString(), NumberStyles.None, CultureInfo.InvariantCulture);
            absValue = decimalExponent >= 0
                ? RatioToDouble(integer * PowerOfTen(decimalExponent), BigInteger.One)
                : RatioToDouble(integer, PowerOfTen(-decimalExponent));
            if (double.IsInfinity(absValue))
                return false;
        }

        value = isNegative ? -absValue : absValue;
        return true;
    }

    private static ulong ParseSmallInteger(StringBuilder digits)
    {
        var integer = 0UL;
        for (var i = 0; i < digits.Length; i++)
            integer = integer * 10 + (ulong)(digits[i] - '0');

        return integer;
    }

    private static bool TryParseExponent(string text, ref int position, int end, out int exponent)
    {
        exponent = 0;
        var isNegative = false;
        if (position < end && (text[position] == '+' || text[position] == '-'))
        {
            isNegative = text[position] == '-';
            position++;
        }

        var start = position;
        for (; position < end && text[position] >= '0' && text[position] <= '9'; position++)
        {
            // A larger exponent gives zero or an overflow, so the method limits it.
            if (exponent < 100000)
                exponent = exponent * 10 + (text[position] - '0');
        }

        if (isNegative)
            exponent = -exponent;

        return position > start;
    }

    private static bool IsWhite(char c)
    {
        // The white space of NumberStyles.AllowLeadingWhite and NumberStyles.AllowTrailingWhite.
        return c == ' ' || (c >= '\t' && c <= '\r');
    }

    /// <summary>
    /// Gets a positive finite value as <paramref name="numerator"/> / <paramref name="denominator"/>.
    /// </summary>
    private static void GetRatio(double absValue, out BigInteger numerator, out BigInteger denominator)
    {
        var bits = BitConverter.DoubleToInt64Bits(absValue);
        var biasedExponent = (int)(bits >> MantissaBits);
        var mantissa = bits & ((1L << MantissaBits) - 1);
        int binaryExponent;
        if (biasedExponent == 0)
        {
            binaryExponent = MinBinaryExponent;
        }
        else
        {
            mantissa |= 1L << MantissaBits;
            binaryExponent = biasedExponent - ExponentBias;
        }

        if (binaryExponent >= 0)
        {
            numerator = new BigInteger(mantissa) << binaryExponent;
            denominator = BigInteger.One;
        }
        else
        {
            numerator = new BigInteger(mantissa);
            denominator = BigInteger.One << -binaryExponent;
        }
    }

    /// <summary>
    /// Gets the exponent of the value in scientific notation, the largest k with 10^k &lt;= value.
    /// </summary>
    private static int GetDecimalExponent(double absValue, BigInteger numerator, BigInteger denominator)
    {
        // The logarithm can be wrong by one near a power of ten, so the exact comparisons correct it.
        var exponent = (int)Math.Floor(Math.Log10(absValue));
        while (CompareWithPowerOfTen(numerator, denominator, exponent) < 0)
            exponent--;
        while (CompareWithPowerOfTen(numerator, denominator, exponent + 1) >= 0)
            exponent++;

        return exponent;
    }

    private static int CompareWithPowerOfTen(BigInteger numerator, BigInteger denominator, int exponent)
    {
        return exponent >= 0
            ? numerator.CompareTo(denominator * PowerOfTen(exponent))
            : (numerator * PowerOfTen(-exponent)).CompareTo(denominator);
    }

    /// <summary>
    /// Rounds the value to <paramref name="precision"/> significant digits. A tie rounds to the even
    /// digit, as the formatter of .NET Core does. If the rounding gives the next power of ten,
    /// the method increments the <paramref name="exponent"/>.
    /// </summary>
    /// <param name="truncatedDigits">The first 17 significant digits of the value, without rounding.</param>
    /// <param name="isRemainderZero">True if the value has no other digits after the 17 digits.</param>
    /// <param name="remainderComparison">The comparison of the other digits with a half unit of the 17th digit.</param>
    /// <param name="precision">The number of significant digits, from 15 to 17.</param>
    /// <param name="exponent">The exponent of the value in scientific notation.</param>
    private static ulong RoundToDigits(ulong truncatedDigits, bool isRemainderZero, int remainderComparison, int precision, ref int exponent)
    {
        ulong digits;
        int comparison;
        if (precision == MaxPrecision)
        {
            digits = truncatedDigits;
            comparison = remainderComparison;
        }
        else
        {
            var divisor = (ulong)ExactPowersOfTen[MaxPrecision - precision];
            digits = truncatedDigits / divisor;
            var droppedDigits = truncatedDigits % divisor;
            var half = divisor / 2;
            comparison = droppedDigits != half
                ? droppedDigits.CompareTo(half)
                : isRemainderZero ? 0 : 1;
        }

        if (comparison > 0 || (comparison == 0 && (digits & 1) != 0))
            digits++;

        if (digits == (ulong)ExactPowersOfTen[precision])
        {
            digits /= 10;
            exponent++;
        }

        return digits;
    }

    private static bool IsRoundTrip(ulong digits, int scale, double absValue)
    {
        double parsed;
        if (digits <= MaxExactInteger && Math.Abs(scale) < ExactPowersOfTen.Length)
        {
            // The same exact calculation as the fast calculation of the parser.
            parsed = scale >= 0
                ? digits * ExactPowersOfTen[scale]
                : digits / ExactPowersOfTen[-scale];
        }
        else
        {
            parsed = scale >= 0
                ? RatioToDouble(new BigInteger(digits) * PowerOfTen(scale), BigInteger.One)
                : RatioToDouble(new BigInteger(digits), PowerOfTen(-scale));
        }

        return parsed.Equals(absValue);
    }

    /// <summary>
    /// Gets the nearest <see cref="double"/> of a positive ratio. A tie rounds to the even mantissa.
    /// </summary>
    /// <returns>The nearest double, or positive infinity if the ratio is too large.</returns>
    internal static double RatioToDouble(BigInteger numerator, BigInteger denominator)
    {
        // The quotient has 54 to 59 bits: 53 bits of the mantissa and the other bits for the rounding.
        // The approximate bit lengths can be wrong by one, so the quotient has a margin.
        var shift = 56 - (GetApproximateBitLength(numerator) - GetApproximateBitLength(denominator));
        var bigQuotient = shift >= 0
            ? BigInteger.DivRem(numerator << shift, denominator, out var remainder)
            : BigInteger.DivRem(numerator, denominator << -shift, out remainder);
        var quotient = (ulong)bigQuotient;

        var droppedBits = GetBitLength(quotient) - (MantissaBits + 1);
        var lowestBitExponent = droppedBits - shift;
        if (lowestBitExponent < MinBinaryExponent)
        {
            // A subnormal value has fewer bits in the mantissa.
            droppedBits += MinBinaryExponent - lowestBitExponent;
            lowestBitExponent = MinBinaryExponent;
        }

        // The value is below half of the smallest subnormal value.
        if (droppedBits >= 64)
            return 0;

        var mantissa = quotient >> droppedBits;
        var droppedValue = quotient & ((1UL << droppedBits) - 1);
        var half = 1UL << (droppedBits - 1);
        if (droppedValue > half || (droppedValue == half && (!remainder.IsZero || (mantissa & 1) != 0)))
            mantissa++;

        if (mantissa == 1UL << (MantissaBits + 1))
        {
            mantissa >>= 1;
            lowestBitExponent++;
        }

        if (mantissa < 1UL << MantissaBits)
        {
            // A subnormal value has the biased exponent zero.
            return BitConverter.Int64BitsToDouble((long)mantissa);
        }

        var biasedExponent = (long)lowestBitExponent + ExponentBias;
        if (biasedExponent >= MaxBiasedExponent)
            return double.PositiveInfinity;

        return BitConverter.Int64BitsToDouble((biasedExponent << MantissaBits) | (long)(mantissa & ((1UL << MantissaBits) - 1)));
    }

    private static int GetApproximateBitLength(BigInteger value)
    {
        return (int)Math.Floor(BigInteger.Log(value, 2)) + 1;
    }

    private static int GetBitLength(ulong value)
    {
        var bitLength = 0;
        for (; value != 0; value >>= 1)
            bitLength++;

        return bitLength;
    }

    private static BigInteger PowerOfTen(int exponent)
    {
        return exponent < PowersOfTen.Length ? PowersOfTen[exponent] : BigInteger.Pow(10, exponent);
    }

    private static BigInteger[] CreatePowersOfTen(int maxExponent)
    {
        var powers = new BigInteger[maxExponent + 1];
        powers[0] = BigInteger.One;
        for (var i = 1; i < powers.Length; i++)
            powers[i] = powers[i - 1] * 10;

        return powers;
    }

    /// <summary>
    /// Formats the digits with the rules of the <c>G</c> format of .NET: fixed-point notation if
    /// the exponent is more than -5 and less than the precision, otherwise scientific notation.
    /// </summary>
    private static string FormatDigits(bool isNegative, ulong digits, int exponent, int precision)
    {
        var digitText = digits.ToString(CultureInfo.InvariantCulture).TrimEnd('0');
        var text = new StringBuilder(32);
        if (isNegative)
            text.Append('-');

        if (exponent > -5 && exponent < precision)
        {
            if (exponent < 0)
            {
                text.Append("0.").Append('0', -exponent - 1).Append(digitText);
            }
            else if (digitText.Length <= exponent + 1)
            {
                text.Append(digitText).Append('0', exponent + 1 - digitText.Length);
            }
            else
            {
                text.Append(digitText, 0, exponent + 1).Append('.').Append(digitText, exponent + 1, digitText.Length - exponent - 1);
            }
        }
        else
        {
            text.Append(digitText[0]);
            if (digitText.Length > 1)
                text.Append('.').Append(digitText, 1, digitText.Length - 1);

            text.Append(exponent < 0 ? "E-" : "E+").Append(Math.Abs(exponent).ToString("00", CultureInfo.InvariantCulture));
        }

        return text.ToString();
    }
}
