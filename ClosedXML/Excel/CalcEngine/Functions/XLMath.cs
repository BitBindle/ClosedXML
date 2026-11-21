using System;
using System.Globalization;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class XLMath
    {
        public static double Asinh(double x)
        {
            return (Math.Log(x + Math.Sqrt(x * x + 1.0)));
        }

        public static double ACosh(double x)
        {
            return (Math.Log(x + Math.Sqrt((x * x) - 1.0)));
        }

        public static double ATanh(double x)
        {
            return (Math.Log((1.0 + x) / (1.0 - x)) / 2.0);
        }

        public static double Csch(double x)
        {
            return (1.0 / Math.Sinh(x));
        }

        internal static OneOf<double, XLError> CombinChecked(CalcContext ctx, double number, double numberChosen)
        {
            if (number < 0 || numberChosen < 0)
                return XLError.NumberInvalid;

            var n = Math.Floor(number);
            var k = Math.Floor(numberChosen);

            // Parameter doesn't fit into int. That's how many multiplications Excel allows.
            if (n >= int.MaxValue || k >= int.MaxValue)
                return XLError.NumberInvalid;

            if (n < k)
                return XLError.NumberInvalid;

            var combinations = Combin(ctx, n, k);
            if (double.IsInfinity(combinations) || double.IsNaN(combinations))
                return XLError.NumberInvalid;

            return combinations;
        }

        internal static double Combin(CalcContext ctx, double n, double k)
        {
            if (k == 0) return 1;

            // Don't use recursion, malicious input could exhaust stack.
            // Don't calculate directly from factorials, could overflow.
            double result = 1;
            for (var i = 1; i <= k; i++, n--)
            {
                ctx.ThrowIfCancelled();
                result *= n;
                result /= i;
            }

            return result;
        }

        internal static double Factorial(double n)
        {
            n = Math.Truncate(n);
            var factorial = 1d;
            while (n > 1)
            {
                factorial *= n--;

                // n can be very large, stop when we reach infinity.
                if (double.IsInfinity(factorial))
                    return factorial;
            }

            return factorial;
        }

        public static Boolean IsEven(Int32 value)
        {
            return Math.Abs(value % 2) == 0;
        }

        public static Boolean IsEven(double value)
        {
            // Check the number doesn't have any fractions and that it is even.
            // Due to rounding after division, only checking for % 2 could fail
            // for numbers really close to whole number.
            var hasNoFraction = value % 1 == 0;
            var isEven = value % 2 == 0;
            return hasNoFraction && isEven;
        }

        public static Boolean IsOdd(Int32 value)
        {
            return Math.Abs(value % 2) != 0;
        }

        public static Boolean IsOdd(double value)
        {
            var hasNoFraction = value % 1 == 0;
            var isOdd = value % 2 != 0;
            return hasNoFraction && isOdd;
        }

        /// <summary>
        /// Compare two numbers at 15 significant digits, like Excel. E.g. <c>0.1+0.2</c> is equal to <c>0.3</c>.
        /// </summary>
        public static int CompareAt15Digits(double left, double right)
        {
            if (left == right)
                return 0;

            // Rounding never changes the order, so numbers that differ well above 15 digits need no rounding.
            if (Math.Abs(left - right) > 1e-13 * Math.Max(Math.Abs(left), Math.Abs(right)))
                return left.CompareTo(right);

            return RoundTo15Digits(left).CompareTo(RoundTo15Digits(right));
        }

        private static double RoundTo15Digits(double value)
        {
            return double.Parse(value.ToString("G15", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        }

        public static double Round(double value, double digits)
        {
            digits = Math.Truncate(digits);

            // Excel rounds the number with 15 significant digits, so ROUND(1.005,2) is 1.01, although
            // the double is 1.00499999999999989. A cast to decimal keeps 15 significant digits,
            // so larger values use the double path to keep their integer digits.
            if (Math.Abs(value) < 1e15 && digits >= -27 && digits <= 27)
            {
                var number = (decimal)value;
                decimal rounded;
                if (digits >= 0)
                {
                    rounded = Math.Round(number, (int)digits, MidpointRounding.AwayFromZero);
                }
                else
                {
                    var coef = (decimal)Math.Pow(10, -digits);
                    rounded = Math.Round(number / coef, MidpointRounding.AwayFromZero) * coef;
                }

                // Decimal keeps the sign of zero, e.g. ROUND(-0.001,2).
                return rounded == 0 ? 0 : (double)rounded;
            }

            if (digits < 0)
            {
                var coef = Math.Pow(10, Math.Abs(digits));
                var shifted = value / coef;
                shifted = Math.Round(shifted, 0, MidpointRounding.AwayFromZero);

                // if coef is infinity
                if (shifted == 0)
                    return 0;

                return shifted * coef;
            }

            // Double can store at most 15 digits and anything below that is float artefact
            return Math.Round(value, (int)Math.Min(digits, 15), MidpointRounding.AwayFromZero);
        }
    }
}
