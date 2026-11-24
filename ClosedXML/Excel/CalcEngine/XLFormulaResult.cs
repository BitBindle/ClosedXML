using System;
using System.Collections.Generic;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Excel
{
    /// <summary>
    /// The result of a formula without implicit intersection. A reference stays a reference, and an array stays
    /// an array.
    /// </summary>
    /// <seealso cref="IXLWorksheet.EvaluateFormula"/>
    public sealed class XLFormulaResult
    {
        private readonly XLCellValue _value;
        private readonly IReadOnlyList<IXLRange>? _ranges;
        private readonly XLCellValue[,]? _array;

        private XLFormulaResult(XLFormulaResultType type, XLCellValue value, IReadOnlyList<IXLRange>? ranges, XLCellValue[,]? array)
        {
            Type = type;
            _value = value;
            _ranges = ranges;
            _array = array;
        }

        /// <summary>
        /// The type of the result. It tells which of <see cref="Value"/>, <see cref="Ranges"/> and
        /// <see cref="Array"/> has the result.
        /// </summary>
        public XLFormulaResultType Type { get; }

        /// <summary>
        /// The value of a result of type <see cref="XLFormulaResultType.Value"/>. A blank result stays blank.
        /// </summary>
        /// <exception cref="InvalidOperationException">The result is a reference or an array.</exception>
        public XLCellValue Value => Type == XLFormulaResultType.Value
            ? _value
            : throw new InvalidOperationException($"The result is not a value. It is {Type}.");

        /// <summary>
        /// The ranges of a result of type <see cref="XLFormulaResultType.Reference"/>, in the order of the
        /// formula. A union, for example <c>(A1:A2,C1)</c>, has more than one range. A reference to more than one
        /// sheet, for example <c>Sheet1:Sheet3!A1</c>, has one range for each sheet in the order of the sheets. Each
        /// range has its worksheet.
        /// </summary>
        /// <exception cref="InvalidOperationException">The result is a value or an array.</exception>
        public IReadOnlyList<IXLRange> Ranges => Type == XLFormulaResultType.Reference
            ? _ranges!
            : throw new InvalidOperationException($"The result is not a reference. It is {Type}.");

        /// <summary>
        /// The values of a result of type <see cref="XLFormulaResultType.Array"/>. The first index is the row and
        /// the second index is the column. A blank element stays blank.
        /// </summary>
        /// <exception cref="InvalidOperationException">The result is a value or a reference.</exception>
        public XLCellValue[,] Array => Type == XLFormulaResultType.Array
            ? _array!
            : throw new InvalidOperationException($"The result is not an array. It is {Type}.");

        /// <summary>
        /// Create the result from a value of the calculation engine.
        /// </summary>
        /// <param name="value">The value of the formula.</param>
        /// <param name="worksheet">The worksheet of an area without a worksheet.</param>
        internal static XLFormulaResult From(AnyValue value, XLWorksheet worksheet)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return new XLFormulaResult(XLFormulaResultType.Value, ToValue(scalar), null, null);

            if (collection.TryPickT0(out var array, out var reference))
            {
                var values = new XLCellValue[array.Height, array.Width];
                for (var row = 0; row < array.Height; row++)
                {
                    for (var column = 0; column < array.Width; column++)
                        values[row, column] = ToValue(array[row, column]);
                }

                return new XLFormulaResult(XLFormulaResultType.Array, default, null, values);
            }

            var ranges = new List<IXLRange>(reference.Areas.Count);
            foreach (var area in reference.Areas)
            {
                var sheet = area.Worksheet ?? worksheet;
                ranges.Add(sheet.Range(in area));
            }

            return new XLFormulaResult(XLFormulaResultType.Reference, default, ranges, null);
        }

        private static XLCellValue ToValue(ScalarValue value)
        {
            return value.Match<XLCellValue>(
                () => Blank.Value,
                logical => logical,
                number => number,
                text => text,
                error => error);
        }
    }
}
