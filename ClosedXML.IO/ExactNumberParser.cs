using System;
using System.Globalization;
using System.Numerics;
using System.Text;

namespace ClosedXML.IO;

/// <summary>
/// Parses a decimal text to the nearest <see cref="double"/>. The result is the IEEE 754 result on each runtime.
/// </summary>
/// <remarks>
/// <para>
/// The <see cref="double"/> parser of .NET Framework is not exact for texts with more than 15
/// significant digits. .NET Core 3.0 and later are exact. The <c>netstandard2.0</c> build reads
/// number texts of a file with this class, so that it reads the same values on each runtime.
/// </para>
/// <para>
/// A text with at most 19 significant digits uses 64-bit and 128-bit integer arithmetic without
/// allocations. Only texts with more digits use <see cref="BigInteger"/>.
/// </para>
/// </remarks>
internal static class ExactNumberParser
{
    private const int MantissaBits = 52;
    private const int ExponentBias = 1075;
    private const int MinBinaryExponent = -1074;
    private const int MaxBiasedExponent = 0x7FF;

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
    /// The smallest decimal exponent of <see cref="PowersOfFive128"/>. 10^-343 * 2^64 is less than half
    /// of the smallest subnormal value.
    /// </summary>
    private const int MinPowerOfTen = -342;

    /// <summary>
    /// The largest decimal exponent of <see cref="PowersOfFive128"/>. 10^309 is too large for a <see cref="double"/>.
    /// </summary>
    private const int MaxPowerOfTen = 308;

    /// <summary>
    /// The number format of <see cref="TryParse"/>: a sign, a decimal point, an exponent and white space.
    /// </summary>
    private const NumberStyles NumberStyle = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign |
                                             NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite |
                                             NumberStyles.AllowExponent;

    /// <summary>
    /// The powers of ten that a <see cref="double"/> keeps exactly.
    /// </summary>
    private static readonly double[] ExactPowersOfTen =
    {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11,
        1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19, 1e20, 1e21, 1e22,
    };

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

        return double.TryParse(text, NumberStyle, CultureInfo.InvariantCulture, out value);
    }

    /// <summary>
    /// Parses a decimal text in the format of <see cref="TryParse"/>, but does not use the parser of the runtime.
    /// </summary>
    /// <returns>
    /// True, if the method found the nearest double. False, if the text has a different format or the
    /// value is too large for a <see cref="double"/>. The caller then decides the result.
    /// </returns>
    internal static bool TryParseExact(string text, out double value)
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
    /// <param name="exponent">The decimal exponent. It is at most <see cref="MaxPowerOfTen"/>.</param>
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

        return MultiplyByPowerOfTen(integer, exponent);
    }

    /// <summary>
    /// Gets the nearest <see cref="double"/> of <paramref name="integer"/> * 10^<paramref name="exponent"/>
    /// with the Eisel-Lemire method. The method multiplies the integer by a 128-bit approximation of
    /// 5^<paramref name="exponent"/>. The fast_float library by Daniel Lemire uses the same calculation.
    /// </summary>
    /// <remarks>
    /// The integer has at most 19 digits. For such an integer, the 128-bit approximation always gives
    /// the nearest double. Noble Mushtak and Daniel Lemire prove this in "Fast number parsing without fallback".
    /// </remarks>
    private static double MultiplyByPowerOfTen(ulong integer, int exponent)
    {
        // The integer is less than 2^64. Thus the value is less than half of the smallest subnormal value.
        if (exponent < MinPowerOfTen)
            return 0;

        var leadingZeros = 64 - GetBitLength(integer);
        var normalizedInteger = integer << leadingZeros;
        var powers = PowersOfFive128.Values;
        var index = 2 * (exponent - MinPowerOfTen);
        Multiply(normalizedInteger, powers[index], out var high, out var low);

        // The result uses the 55 high bits of the product. If the other bits of the high word are all
        // ones, the low word of the power can change these bits.
        const ulong precisionMask = ulong.MaxValue >> (MantissaBits + 3);
        if ((high & precisionMask) == precisionMask)
        {
            Multiply(normalizedInteger, powers[index + 1], out var secondHigh, out _);
            low += secondHigh;
            if (secondHigh > low)
                high++;
        }

        var upperBit = (int)(high >> 63);
        var shift = upperBit + 64 - MantissaBits - 3;
        var mantissa = high >> shift;

        // The binary exponent of 10^exponent is floor(exponent * log2(10)).
        var binaryExponent = ((217706 * exponent) >> 16) + 63;
        var biasedExponent = binaryExponent + upperBit - leadingZeros + 1023;
        if (biasedExponent <= 0)
        {
            // A subnormal value has fewer bits in the mantissa.
            if (-biasedExponent + 1 >= 64)
                return 0;

            mantissa >>= -biasedExponent + 1;
            mantissa += mantissa & 1;
            mantissa >>= 1;
            return BitConverter.Int64BitsToDouble((long)mantissa);
        }

        // An exact tie has zero bits below the rounding bit. Only these exponents can give an exact tie.
        // The tie selects the even mantissa.
        if (low <= 1 && exponent >= -4 && exponent <= 23 && (mantissa & 3) == 1 && mantissa << shift == high)
            mantissa &= ~1UL;

        mantissa += mantissa & 1;
        mantissa >>= 1;
        if (mantissa >= 2UL << MantissaBits)
        {
            mantissa >>= 1;
            biasedExponent++;
        }

        if (biasedExponent >= MaxBiasedExponent)
            return double.PositiveInfinity;

        return BitConverter.Int64BitsToDouble(((long)biasedExponent << MantissaBits) | (long)(mantissa & ((1UL << MantissaBits) - 1)));
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

    private static int GetBitLength(ulong value)
    {
        var length = 0;
        if (value >= 1UL << 32)
        {
            value >>= 32;
            length += 32;
        }

        if (value >= 1UL << 16)
        {
            value >>= 16;
            length += 16;
        }

        if (value >= 1UL << 8)
        {
            value >>= 8;
            length += 8;
        }

        if (value >= 1UL << 4)
        {
            value >>= 4;
            length += 4;
        }

        if (value >= 1UL << 2)
        {
            value >>= 2;
            length += 2;
        }

        if (value >= 1UL << 1)
        {
            value >>= 1;
            length++;
        }

        return length + (int)value;
    }

    /// <summary>
    /// The powers of ten of the <see cref="BigInteger"/> calculations. The runtime creates the cache
    /// only when a value needs it, because most files have no such values.
    /// </summary>
    private static class BigPowersOfTen
    {
        private static readonly BigInteger[] Cache = Create(350);

        /// <summary>
        /// A static constructor stops the runtime from creating the cache before the first use.
        /// </summary>
        static BigPowersOfTen()
        {
        }

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

    /// <summary>
    /// The 128-bit approximations of the powers of five from 5^<see cref="MinPowerOfTen"/> to
    /// 5^<see cref="MaxPowerOfTen"/>. Each power has two items: the high word and the low word.
    /// The runtime creates the table only when a value needs it.
    /// </summary>
    /// <remarks>
    /// Each approximation has 128 bits, and its highest bit is set. The table has the same values
    /// as the table of the fast_float library, because the error limits of the method use them.
    /// </remarks>
    internal static class PowersOfFive128
    {
        internal static readonly ulong[] Values = Create();

        /// <summary>
        /// A static constructor stops the runtime from creating the table before the first use.
        /// </summary>
        static PowersOfFive128()
        {
        }

        private static ulong[] Create()
        {
            var values = new ulong[2 * (MaxPowerOfTen - MinPowerOfTen + 1)];
            var power = BigInteger.Pow(5, -MinPowerOfTen);
            for (var exponent = MinPowerOfTen; exponent <= MaxPowerOfTen; exponent++)
            {
                BigInteger approximation;
                if (exponent < 0)
                {
                    // A negative exponent uses the reciprocal 2^scale / 5^-exponent, rounded up.
                    var bitLength = GetBitLength(power);
                    var scale = exponent >= -27 ? bitLength + 127 : 2 * bitLength + 128;
                    approximation = (BigInteger.One << scale) / power + 1;
                    power /= 5;
                }
                else
                {
                    approximation = power;
                    power *= 5;
                }

                // Truncate the approximation to 128 bits, or move its highest bit to bit 127.
                var shift = GetBitLength(approximation) - 128;
                approximation = shift >= 0 ? approximation >> shift : approximation << -shift;
                var index = 2 * (exponent - MinPowerOfTen);
                values[index] = (ulong)(approximation >> 64);
                values[index + 1] = (ulong)(approximation & ulong.MaxValue);
            }

            return values;
        }

        private static int GetBitLength(BigInteger value)
        {
            var bytes = value.ToByteArray();
            var highByte = bytes[bytes.Length - 1];
            return (bytes.Length - 1) * 8 + ExactNumberParser.GetBitLength(highByte);
        }
    }
}
