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
/// <para>
/// The <see cref="double"/> parser and formatter of .NET Framework are not exact for texts
/// with more than 15 significant digits. .NET Core 3.0 and later are exact. The netstandard2.0
/// build uses this class, so that it reads and writes the same values on each runtime.
/// </para>
/// <para>
/// The reader and the writer call this class for each number cell. The usual values use 64-bit
/// and 128-bit integer arithmetic without allocations. Only very small, very large and very long
/// values use <see cref="BigInteger"/>.
/// </para>
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
    /// An unsigned long keeps each integer with 19 digits.
    /// </summary>
    private const int MaxSmallIntegerDigits = 19;

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
    /// The nearest doubles of the powers of ten that have a power of five in <see cref="PowersOfFive"/>.
    /// The 128-bit calculation uses them only for its first estimate.
    /// </summary>
    private static readonly double[] ApproximatePowersOfTen =
    {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13,
        1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22, 1e23, 1e24, 1e25, 1e26, 1e27,
    };

    /// <summary>
    /// The powers of five that an unsigned long keeps. 5^27 is the largest.
    /// </summary>
    private static readonly ulong[] PowersOfFive = CreatePowers(5, 27);

    /// <summary>
    /// The powers of ten that an unsigned long keeps. 10^19 is the largest.
    /// </summary>
    private static readonly ulong[] IntegerPowersOfTen = CreatePowers(10, 19);

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

        GetBinary(absValue, out var mantissa, out var binaryExponent);
        if (!TryGetTruncatedDigits(absValue, mantissa, binaryExponent, out var truncatedDigits, out var valueExponent, out var isRemainderZero, out var remainderComparison))
            GetTruncatedDigitsOfAnyValue(absValue, mantissa, binaryExponent, out truncatedDigits, out valueExponent, out isRemainderZero, out remainderComparison);

        for (var precision = 15; ; precision++)
        {
            var exponent = valueExponent;
            var digits = RoundToDigits(truncatedDigits, isRemainderZero, remainderComparison, precision, ref exponent);

            // 17 significant digits are sufficient for each double value.
            if (precision == MaxPrecision || ToDouble(digits, exponent - precision + 1).Equals(absValue))
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

        return double.TryParse(text, XLHelper.NumberStyle, XLHelper.ParseCulture, out value);
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

        // The integer keeps the first 19 significant digits. The digit count includes the other digits.
        var digitsStart = position;
        var integer = 0UL;
        var digitCount = 0;
        var hasDigit = false;
        var hasDecimalPoint = false;
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
            if (digitCount == 0 && c == '0')
            {
                if (hasDecimalPoint)
                    decimalExponent--;

                continue;
            }

            if (digitCount < MaxSmallIntegerDigits)
            {
                integer = integer * 10 + (ulong)(c - '0');
                if (hasDecimalPoint)
                    decimalExponent--;
            }
            else if (!hasDecimalPoint)
            {
                decimalExponent++;
            }

            digitCount++;
        }

        var digitsEnd = position;
        if (!hasDigit)
            return false;

        var exponentPart = 0;
        if (position < end && (text[position] == 'e' || text[position] == 'E'))
        {
            position++;
            if (!TryParseExponent(text, ref position, end, out exponentPart))
                return false;
        }

        if (position != end)
            return false;

        double absValue;
        if (digitCount > MaxSmallIntegerDigits)
        {
            if (!TryParseManyDigits(text, digitsStart, digitsEnd, exponentPart, out absValue))
                return false;
        }
        else if (digitCount == 0)
        {
            absValue = 0;
        }
        else
        {
            // Without the trailing zeros, more texts can use the fast calculations.
            while (integer % 10 == 0)
            {
                integer /= 10;
                digitCount--;
                decimalExponent++;
            }

            var scale = decimalExponent + exponentPart;
            if (digitCount + scale < -330)
            {
                absValue = 0;
            }
            else
            {
                // The value is 10^309 or more. The runtime decides the result of an overflow.
                if (digitCount + scale > 309)
                    return false;

                absValue = ToDouble(integer, scale);
                if (double.IsInfinity(absValue))
                    return false;
            }
        }

        value = isNegative ? -absValue : absValue;
        return true;
    }

    /// <summary>
    /// Parses the digits of a text that has more than 19 significant digits.
    /// </summary>
    private static bool TryParseManyDigits(string text, int start, int end, int exponentPart, out double absValue)
    {
        absValue = 0;
        var digits = new StringBuilder();
        var decimalExponent = exponentPart;
        var hasDecimalPoint = false;
        var hasDiscardedNonZeroDigit = false;
        for (var i = start; i < end; i++)
        {
            var c = text[i];
            if (c == '.')
            {
                hasDecimalPoint = true;
                continue;
            }

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

        if (hasDiscardedNonZeroDigit)
        {
            // A digit below all kept digits keeps the information that the value is above the kept digits.
            digits.Append('1');
            decimalExponent--;
        }

        if (digits.Length + decimalExponent < -330)
            return true;

        // The value is 10^309 or more. The runtime decides the result of an overflow.
        if (digits.Length + decimalExponent > 309)
            return false;

        var integer = BigInteger.Parse(digits.ToString(), NumberStyles.None, CultureInfo.InvariantCulture);
        absValue = decimalExponent >= 0
            ? RatioToDouble(integer * BigPowersOfTen.Get(decimalExponent), BigInteger.One)
            : RatioToDouble(integer, BigPowersOfTen.Get(-decimalExponent));
        return !double.IsInfinity(absValue);
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
    /// Gets the nearest <see cref="double"/> of <paramref name="integer"/> * 10^<paramref name="exponent"/>.
    /// </summary>
    /// <param name="integer">A positive integer.</param>
    /// <param name="exponent">The decimal exponent.</param>
    private static double ToDouble(ulong integer, int exponent)
    {
        if (integer <= MaxExactInteger && Math.Abs(exponent) < ExactPowersOfTen.Length)
        {
            // The integer and the power of ten are exact doubles, and IEEE 754 rounds the
            // product or the quotient one time. Thus the result is the nearest double.
            return exponent >= 0
                ? (long)integer * ExactPowersOfTen[exponent]
                : (long)integer / ExactPowersOfTen[-exponent];
        }

        if (Math.Abs(exponent) < PowersOfFive.Length)
            return MultiplyByPowerOfTen(integer, exponent);

        return exponent >= 0
            ? RatioToDouble(new BigInteger(integer) * BigPowersOfTen.Get(exponent), BigInteger.One)
            : RatioToDouble(new BigInteger(integer), BigPowersOfTen.Get(-exponent));
    }

    /// <summary>
    /// Gets the nearest <see cref="double"/> of <paramref name="integer"/> * 10^<paramref name="exponent"/>
    /// with 128-bit integer arithmetic. The absolute value of the exponent is at most 27.
    /// </summary>
    private static double MultiplyByPowerOfTen(ulong integer, int exponent)
    {
        // The estimate is at most a few units in the last place from the result.
        // The loop moves it to the nearest double. A tie selects the even mantissa.
        var candidate = integer < 1UL << 63 ? (long)integer : (long)(integer >> 1) * 2.0;
        candidate = exponent >= 0
            ? candidate * ApproximatePowersOfTen[exponent]
            : candidate / ApproximatePowersOfTen[-exponent];
        while (true)
        {
            var comparison = CompareWithMidpointAbove(integer, exponent, candidate);
            if (comparison > 0 || (comparison == 0 && HasOddMantissa(candidate)))
            {
                candidate = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(candidate) + 1);
                continue;
            }

            var below = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(candidate) - 1);
            comparison = CompareWithMidpointAbove(integer, exponent, below);
            if (comparison < 0 || (comparison == 0 && !HasOddMantissa(below)))
            {
                candidate = below;
                continue;
            }

            return candidate;
        }
    }

    private static bool HasOddMantissa(double value)
    {
        return (BitConverter.DoubleToInt64Bits(value) & 1) != 0;
    }

    /// <summary>
    /// Compares <paramref name="integer"/> * 10^<paramref name="exponent"/> with the midpoint between
    /// <paramref name="candidate"/> and the next larger double.
    /// </summary>
    private static int CompareWithMidpointAbove(ulong integer, int exponent, double candidate)
    {
        // The candidate is m * 2^e, the next double is (m + 1) * 2^e and the midpoint is (2m + 1) * 2^(e - 1).
        // A power of ten is 5^n * 2^n. Each side of the comparison gets the power of five that it needs.
        GetBinary(candidate, out var mantissa, out var binaryExponent);
        var midpointMantissa = (mantissa << 1) | 1;
        ulong valueHigh, valueLow, midpointHigh, midpointLow;
        if (exponent >= 0)
        {
            Multiply(integer, PowersOfFive[exponent], out valueHigh, out valueLow);
            midpointHigh = 0;
            midpointLow = midpointMantissa;
        }
        else
        {
            valueHigh = 0;
            valueLow = integer;
            Multiply(midpointMantissa, PowersOfFive[-exponent], out midpointHigh, out midpointLow);
        }

        return CompareScaled(valueHigh, valueLow, exponent, midpointHigh, midpointLow, binaryExponent - 1);
    }

    /// <summary>
    /// Compares x * 2^<paramref name="xExponent"/> with y * 2^<paramref name="yExponent"/>.
    /// x and y are positive 128-bit integers.
    /// </summary>
    private static int CompareScaled(ulong xHigh, ulong xLow, int xExponent, ulong yHigh, ulong yLow, int yExponent)
    {
        var xLength = GetBitLength(xHigh, xLow) + xExponent;
        var yLength = GetBitLength(yHigh, yLow) + yExponent;
        if (xLength != yLength)
            return xLength.CompareTo(yLength);

        // Both values have the same bit length, so each shift keeps the value in 128 bits.
        if (xExponent > yExponent)
            ShiftLeft(ref xHigh, ref xLow, xExponent - yExponent);
        else
            ShiftLeft(ref yHigh, ref yLow, yExponent - xExponent);

        return Compare(xHigh, xLow, yHigh, yLow);
    }

    /// <summary>
    /// Gets a positive finite value as <paramref name="mantissa"/> * 2^<paramref name="binaryExponent"/>.
    /// </summary>
    private static void GetBinary(double absValue, out ulong mantissa, out int binaryExponent)
    {
        var bits = (ulong)BitConverter.DoubleToInt64Bits(absValue);
        var biasedExponent = (int)(bits >> MantissaBits);
        mantissa = bits & ((1UL << MantissaBits) - 1);
        if (biasedExponent == 0)
        {
            binaryExponent = MinBinaryExponent;
        }
        else
        {
            mantissa |= 1UL << MantissaBits;
            binaryExponent = biasedExponent - ExponentBias;
        }
    }

    /// <summary>
    /// Gets the first 17 significant digits of a value from 10^-11 to 10^17 with 128-bit integer arithmetic.
    /// </summary>
    /// <returns>False if the value is outside the range.</returns>
    private static bool TryGetTruncatedDigits(double absValue, ulong mantissa, int binaryExponent, out ulong truncatedDigits, out int exponent, out bool isRemainderZero, out int remainderComparison)
    {
        // The logarithm can be wrong by one near a power of ten, so the loop corrects it.
        exponent = (int)Math.Floor(Math.Log10(absValue));
        while (true)
        {
            truncatedDigits = 0;
            isRemainderZero = true;
            remainderComparison = -1;

            // The digits are value * 10^scale = mantissa * 5^scale * 2^(binaryExponent + scale).
            var scale = MaxPrecision - 1 - exponent;
            if (scale < 0 || scale >= PowersOfFive.Length)
                return false;

            Multiply(mantissa, PowersOfFive[scale], out var high, out var low);
            var shift = binaryExponent + scale;
            if (shift >= 0)
            {
                if (GetBitLength(high, low) + shift > 64)
                {
                    exponent++;
                    continue;
                }

                ShiftLeft(ref high, ref low, shift);
                truncatedDigits = low;
            }
            else
            {
                // The values of the range have a shift from -63 to -1.
                var rightShift = -shift;
                if (rightShift >= 64)
                    return false;

                var digitsHigh = high >> rightShift;
                var digitsLow = (low >> rightShift) | (high << (64 - rightShift));
                if (digitsHigh != 0)
                {
                    exponent++;
                    continue;
                }

                // The remainder is the value minus the digits, in units of 2^shift.
                var truncatedHigh = digitsHigh;
                var truncatedLow = digitsLow;
                ShiftLeft(ref truncatedHigh, ref truncatedLow, rightShift);
                var remainderLow = low - truncatedLow;
                var remainderHigh = high - truncatedHigh - (low < truncatedLow ? 1UL : 0UL);
                var halfHigh = 0UL;
                var halfLow = 1UL;
                ShiftLeft(ref halfHigh, ref halfLow, rightShift - 1);

                truncatedDigits = digitsLow;
                isRemainderZero = remainderHigh == 0 && remainderLow == 0;
                remainderComparison = Compare(remainderHigh, remainderLow, halfHigh, halfLow);
            }

            if (truncatedDigits >= IntegerPowersOfTen[MaxPrecision])
            {
                exponent++;
                continue;
            }

            if (truncatedDigits < IntegerPowersOfTen[MaxPrecision - 1])
            {
                exponent--;
                continue;
            }

            return true;
        }
    }

    /// <summary>
    /// Gets the first 17 significant digits of each positive finite value with <see cref="BigInteger"/>.
    /// </summary>
    private static void GetTruncatedDigitsOfAnyValue(double absValue, ulong mantissa, int binaryExponent, out ulong truncatedDigits, out int exponent, out bool isRemainderZero, out int remainderComparison)
    {
        var numerator = binaryExponent >= 0 ? new BigInteger(mantissa) << binaryExponent : new BigInteger(mantissa);
        var denominator = binaryExponent >= 0 ? BigInteger.One : BigInteger.One << -binaryExponent;

        // The logarithm can be wrong by one near a power of ten, so the loop corrects it.
        exponent = (int)Math.Floor(Math.Log10(absValue));
        while (true)
        {
            var scale = MaxPrecision - 1 - exponent;
            var scaledNumerator = scale < 0 ? numerator : numerator * BigPowersOfTen.Get(scale);
            var scaledDenominator = scale < 0 ? denominator * BigPowersOfTen.Get(-scale) : denominator;
            var digits = BigInteger.DivRem(scaledNumerator, scaledDenominator, out var remainder);
            if (digits >= IntegerPowersOfTen[MaxPrecision])
            {
                exponent++;
                continue;
            }

            if (digits < IntegerPowersOfTen[MaxPrecision - 1])
            {
                exponent--;
                continue;
            }

            truncatedDigits = (ulong)digits;
            isRemainderZero = remainder.IsZero;
            remainderComparison = (remainder << 1).CompareTo(scaledDenominator);
            return;
        }
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
            var divisor = IntegerPowersOfTen[MaxPrecision - precision];
            digits = truncatedDigits / divisor;
            var droppedDigits = truncatedDigits % divisor;
            var half = divisor / 2;
            comparison = droppedDigits != half
                ? droppedDigits.CompareTo(half)
                : isRemainderZero ? 0 : 1;
        }

        if (comparison > 0 || (comparison == 0 && (digits & 1) != 0))
            digits++;

        if (digits == IntegerPowersOfTen[precision])
        {
            digits /= 10;
            exponent++;
        }

        return digits;
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

        var droppedBits = GetBitLength(0, quotient) - (MantissaBits + 1);
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

    private static void Multiply(ulong x, ulong y, out ulong high, out ulong low)
    {
        var xLow = (ulong)(uint)x;
        var xHigh = x >> 32;
        var yLow = (ulong)(uint)y;
        var yHigh = y >> 32;
        var lowLow = xLow * yLow;
        var lowHigh = xLow * yHigh;
        var highLow = xHigh * yLow;
        var middle = (lowLow >> 32) + (uint)lowHigh + (uint)highLow;
        low = (middle << 32) | (uint)lowLow;
        high = (xHigh * yHigh) + (lowHigh >> 32) + (highLow >> 32) + (middle >> 32);
    }

    private static void ShiftLeft(ref ulong high, ref ulong low, int count)
    {
        if (count == 0)
            return;

        if (count >= 64)
        {
            high = low << (count - 64);
            low = 0;
        }
        else
        {
            high = (high << count) | (low >> (64 - count));
            low <<= count;
        }
    }

    private static int Compare(ulong xHigh, ulong xLow, ulong yHigh, ulong yLow)
    {
        return xHigh != yHigh ? xHigh.CompareTo(yHigh) : xLow.CompareTo(yLow);
    }

    private static int GetBitLength(ulong high, ulong low)
    {
        if (high != 0)
            return 64 + GetBitLength(0, high);

        var length = 0;
        if (low >= 1UL << 32)
        {
            low >>= 32;
            length += 32;
        }

        if (low >= 1UL << 16)
        {
            low >>= 16;
            length += 16;
        }

        if (low >= 1UL << 8)
        {
            low >>= 8;
            length += 8;
        }

        if (low >= 1UL << 4)
        {
            low >>= 4;
            length += 4;
        }

        if (low >= 1UL << 2)
        {
            low >>= 2;
            length += 2;
        }

        if (low >= 1UL << 1)
        {
            low >>= 1;
            length++;
        }

        return length + (int)low;
    }

    private static ulong[] CreatePowers(ulong radix, int maxExponent)
    {
        var powers = new ulong[maxExponent + 1];
        powers[0] = 1;
        for (var i = 1; i < powers.Length; i++)
            powers[i] = powers[i - 1] * radix;

        return powers;
    }

    /// <summary>
    /// Formats the digits with the rules of the <c>G</c> format of .NET: fixed-point notation if
    /// the exponent is more than -5 and less than the precision, otherwise scientific notation.
    /// </summary>
    private static string FormatDigits(bool isNegative, ulong digits, int exponent, int precision)
    {
        while (digits % 10 == 0)
            digits /= 10;

        var digitCount = 1;
        while (digitCount < MaxPrecision && digits >= IntegerPowersOfTen[digitCount])
            digitCount++;

        // The longest text, for example "-1.2345678901234567E-308", has 24 characters.
        Span<char> digitText = stackalloc char[MaxPrecision];
        for (var i = digitCount - 1; i >= 0; i--)
        {
            digitText[i] = (char)('0' + (int)(digits % 10));
            digits /= 10;
        }

        Span<char> text = stackalloc char[32];
        var length = 0;
        if (isNegative)
            text[length++] = '-';

        if (exponent > -5 && exponent < precision)
        {
            if (exponent < 0)
            {
                text[length++] = '0';
                text[length++] = '.';
                for (var i = 0; i < -exponent - 1; i++)
                    text[length++] = '0';

                Append(text, ref length, digitText.Slice(0, digitCount));
            }
            else if (digitCount <= exponent + 1)
            {
                Append(text, ref length, digitText.Slice(0, digitCount));
                for (var i = digitCount; i < exponent + 1; i++)
                    text[length++] = '0';
            }
            else
            {
                Append(text, ref length, digitText.Slice(0, exponent + 1));
                text[length++] = '.';
                Append(text, ref length, digitText.Slice(exponent + 1, digitCount - exponent - 1));
            }
        }
        else
        {
            text[length++] = digitText[0];
            if (digitCount > 1)
            {
                text[length++] = '.';
                Append(text, ref length, digitText.Slice(1, digitCount - 1));
            }

            text[length++] = 'E';
            text[length++] = exponent < 0 ? '-' : '+';
            var absExponent = Math.Abs(exponent);
            if (absExponent >= 100)
                text[length++] = (char)('0' + absExponent / 100);

            text[length++] = (char)('0' + absExponent / 10 % 10);
            text[length++] = (char)('0' + absExponent % 10);
        }

        return text.Slice(0, length).ToString();
    }

    private static void Append(Span<char> text, ref int length, ReadOnlySpan<char> value)
    {
        value.CopyTo(text.Slice(length));
        length += value.Length;
    }

    /// <summary>
    /// The powers of ten of the <see cref="BigInteger"/> calculations. The runtime creates the cache
    /// only when a value needs it, because most files have no such values.
    /// </summary>
    private static class BigPowersOfTen
    {
        private static readonly BigInteger[] Cache = Create(350);

        internal static BigInteger Get(int exponent)
        {
            return exponent < Cache.Length ? Cache[exponent] : BigInteger.Pow(10, exponent);
        }

        private static BigInteger[] Create(int maxExponent)
        {
            var powers = new BigInteger[maxExponent + 1];
            powers[0] = BigInteger.One;
            for (var i = 1; i < powers.Length; i++)
                powers[i] = powers[i - 1] * 10;

            return powers;
        }
    }
}
