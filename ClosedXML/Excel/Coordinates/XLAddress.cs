using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace ClosedXML.Excel
{
    internal struct XLAddress : IXLAddress, IEquatable<XLAddress>
    {
        #region Static
        /// <summary>
        /// Create address without worksheet. For calculation only!
        /// </summary>
        /// <param name="cellAddressString"></param>
        public static XLAddress Create(string cellAddressString)
        {
            return Create(null, cellAddressString);
        }

        public static XLAddress Create(XLWorksheet? worksheet, string cellAddressString)
        {
            var fixedColumn = cellAddressString[0] == '$';
            Int32 startPos;
            if (fixedColumn)
            {
                startPos = 1;
            }
            else
            {
                startPos = 0;
            }

            int rowPos = startPos;
            while (cellAddressString[rowPos] > '9')
            {
                rowPos++;
            }

            var fixedRow = cellAddressString[rowPos] == '$';
            string columnLetter;
            int rowNumber;
            if (fixedRow)
            {
                if (fixedColumn)
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
                }
                else
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos);
                }

                rowNumber = int.Parse(cellAddressString.Substring(rowPos + 1), XLHelper.NumberStyle, XLHelper.ParseCulture);
            }
            else
            {
                if (fixedColumn)
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
                }
                else
                {
                    columnLetter = cellAddressString.Substring(startPos, rowPos);
                }

                rowNumber = Int32.Parse(cellAddressString.Substring(rowPos), XLHelper.NumberStyle, XLHelper.ParseCulture);
            }
            return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
        }

        #endregion Static

        #region Private fields

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private bool _fixedRow;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private bool _fixedColumn;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int _rowNumber;

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private readonly int _columnNumber;

        private string? _trimmedAddress;

        /// <summary>
        /// The length of the longest A1 text of a valid address, <c>$XFD$1048576</c>.
        /// </summary>
        internal const int MaxA1Length = 12;

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.  Attention: without worksheet for calculation only!
        /// </summary>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnLetter">The column letter of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
                : this(null, rowNumber, columnLetter, fixedRow, fixedColumn)
        {
        }

        /// <summary>
        /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnLetter">The column letter of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet? worksheet, int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
                : this(worksheet, rowNumber, XLHelper.GetColumnNumberFromLetter(columnLetter), fixedRow, fixedColumn)
        {
        }

        /// <summary>
        /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation. Attention: without worksheet for calculation only!
        /// </summary>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnNumber">The column number of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)
                : this(null, rowNumber, columnNumber, fixedRow, fixedColumn)
        {
        }

        /// <summary>
        /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation.
        /// </summary>
        /// <param name = "worksheet"></param>
        /// <param name = "rowNumber">The row number of the cell address.</param>
        /// <param name = "columnNumber">The column number of the cell address.</param>
        /// <param name = "fixedRow"></param>
        /// <param name = "fixedColumn"></param>
        public XLAddress(XLWorksheet? worksheet, int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn) : this()

        {
            Worksheet = worksheet;

            _rowNumber = rowNumber;
            _columnNumber = columnNumber;
            _fixedColumn = fixedColumn;
            _fixedRow = fixedRow;
        }

        #endregion Constructors

        #region Properties

        public XLWorksheet? Worksheet { get; internal set; }

        IXLWorksheet? IXLAddress.Worksheet
        {
            [DebuggerStepThrough]
            get { return Worksheet; }
        }

        [MemberNotNullWhen(true, nameof(Worksheet))]
        public bool HasWorksheet
        {
            [DebuggerStepThrough]
            get { return Worksheet != null; }
        }

        public bool FixedRow
        {
            get { return _fixedRow; }
        }

        public bool FixedColumn
        {
            get { return _fixedColumn; }
        }

        /// <summary>
        /// Gets the row number of this address.
        /// </summary>
        public Int32 RowNumber
        {
            get { return _rowNumber; }
        }

        /// <summary>
        /// Gets the column number of this address.
        /// </summary>
        public Int32 ColumnNumber
        {
            get { return _columnNumber; }
        }

        /// <summary>
        /// Gets the column letter(s) of this address.
        /// </summary>
        public String ColumnLetter
        {
            get { return XLHelper.GetColumnLetterFromNumber(_columnNumber); }
        }

        #endregion Properties

        #region Overrides

        public override string ToString()
        {
            if (!IsValid)
                return "#REF!";

            return CreateA1Text(_fixedColumn, _fixedRow);
        }

        public string ToString(XLReferenceStyle referenceStyle)
        {
            return ToString(referenceStyle, false);
        }

        public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            string address;
            if (!IsValid)
                address = "#REF!";
            else if (referenceStyle == XLReferenceStyle.A1)
                address = GetTrimmedAddress();
            else if (referenceStyle == XLReferenceStyle.R1C1
                     || HasWorksheet && Worksheet.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
                address = CreateR1C1Text();
            else
                address = GetTrimmedAddress();

            if (includeSheet)
                return String.Concat(
                    GetSheetPrefix(),
                    '!',
                    address);

            return address;
        }

        #endregion Overrides

        #region Methods

        public string GetTrimmedAddress()
        {
            return _trimmedAddress ??= CreateA1Text(fixedColumn: false, fixedRow: false);
        }

        /// <summary>
        /// Write the A1 text of an address with a valid column to the span.
        /// </summary>
        /// <param name="output">The span. It must have space for <see cref="MaxA1Length"/> characters.</param>
        /// <param name="fixedColumn">Write a <c>$</c> before the column letters.</param>
        /// <param name="fixedRow">Write a <c>$</c> before the row number.</param>
        /// <returns>The number of characters that the method writes.</returns>
        internal int FormatA1(Span<char> output, bool fixedColumn, bool fixedRow)
        {
            var length = 0;
            if (fixedColumn)
                output[length++] = '$';

            var columnLetter = ColumnLetter;
            columnLetter.CopyTo(output[length..]);
            length += columnLetter.Length;

            if (fixedRow)
                output[length++] = '$';

            var isWritten = _rowNumber.TryFormat(output[length..], out var rowLength, provider: CultureInfo.InvariantCulture);
            Debug.Assert(isWritten, "The span is too short for the row number.");
            return length + rowLength;
        }

        private string CreateA1Text(bool fixedColumn, bool fixedRow)
        {
            // The row number of an address that is not valid can have more digits.
            Span<char> text = stackalloc char[MaxA1Length + 4];
            var length = FormatA1(text, fixedColumn, fixedRow);
            return text[..length].ToString();
        }

        private string CreateR1C1Text()
        {
            return string.Create(CultureInfo.InvariantCulture, $"R{_rowNumber}C{_columnNumber}");
        }

        #endregion Methods

        #region Operator Overloads

        public static XLAddress operator +(XLAddress left, XLAddress right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right.RowNumber,
                                 left.ColumnNumber + right.ColumnNumber,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, XLAddress right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right.RowNumber,
                                 left.ColumnNumber - right.ColumnNumber,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator +(XLAddress left, Int32 right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber + right,
                                 left.ColumnNumber + right,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static XLAddress operator -(XLAddress left, Int32 right)
        {
            return new XLAddress(left.Worksheet,
                                 left.RowNumber - right,
                                 left.ColumnNumber - right,
                                 left._fixedRow,
                                 left._fixedColumn);
        }

        public static Boolean operator ==(XLAddress left, XLAddress right)
        {
            return left.Equals(right);
        }

        public static Boolean operator !=(XLAddress left, XLAddress right)
        {
            return !(left == right);
        }

        #endregion Operator Overloads

        #region Interface Requirements

        #region IEqualityComparer<XLCellAddress> Members

        public Boolean Equals(IXLAddress? x, IXLAddress? y)
        {
            return x == y;
        }

        public new Boolean Equals(object x, object y)
        {
            return x == y;
        }

        #endregion IEqualityComparer<XLCellAddress> Members

        #region IEquatable<XLCellAddress> Members

        public bool Equals(IXLAddress? other)
        {
            if (other == null)
                return false;

            return _rowNumber == other.RowNumber &&
                   _columnNumber == other.ColumnNumber &&
                   _fixedRow == other.FixedRow &&
                   _fixedColumn == other.FixedColumn;
        }

        public bool Equals(XLAddress other)
        {
            return _rowNumber == other._rowNumber &&
                   _columnNumber == other._columnNumber &&
                   _fixedRow == other._fixedRow &&
                   _fixedColumn == other._fixedColumn;
        }

        public override Boolean Equals(Object? other)
        {
            return Equals(other as IXLAddress);
        }

        public override int GetHashCode()
        {
            var hashCode = 2122234362;
            hashCode = hashCode * -1521134295 + _fixedRow.GetHashCode();
            hashCode = hashCode * -1521134295 + _fixedColumn.GetHashCode();
            hashCode = hashCode * -1521134295 + _rowNumber.GetHashCode();
            hashCode = hashCode * -1521134295 + _columnNumber.GetHashCode();
            return hashCode;
        }

        public int GetHashCode(IXLAddress obj)
        {
            return ((XLAddress)obj).GetHashCode();
        }

        #endregion IEquatable<XLCellAddress> Members

        #endregion Interface Requirements

        public String ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public String ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.Default);
        }

        public String ToStringRelative(Boolean includeSheet)
        {
            var address = IsValid ? GetTrimmedAddress() : "#REF!";

            if (includeSheet)
                return String.Concat(
                    GetSheetPrefix(),
                    '!',
                    address
                );

            return address;
        }

        internal XLAddress WithoutWorksheet()
        {
            return new XLAddress(RowNumber, ColumnNumber, FixedRow, FixedColumn);
        }

        internal XLAddress WithWorksheet(XLWorksheet worksheet)
        {
            return new XLAddress(worksheet, RowNumber, ColumnNumber, FixedRow, FixedColumn);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
        {
            String address;

            if (referenceStyle == XLReferenceStyle.Default && HasWorksheet)
                referenceStyle = Worksheet.Workbook.ReferenceStyle;

            if (referenceStyle == XLReferenceStyle.Default)
                referenceStyle = XLReferenceStyle.A1;

            Debug.Assert(referenceStyle != XLReferenceStyle.Default);

            if (!IsValid)
            {
                address = "#REF!";
            }
            else
            {
                switch (referenceStyle)
                {
                    case XLReferenceStyle.A1:
                        address = CreateA1Text(fixedColumn: true, fixedRow: true);
                        break;

                    case XLReferenceStyle.R1C1:
                        address = CreateR1C1Text();
                        break;

                    default:
                        throw new NotImplementedException();
                }
            }

            if (includeSheet)
                return String.Concat(
                    GetSheetPrefix(),
                    '!',
                    address);

            return address;
        }

        public String UniqueId { get { return RowNumber.ToString("0000000") + ColumnNumber.ToString("00000"); } }

        public bool IsValid
        {
            get
            {
                return 0 < RowNumber && RowNumber <= XLHelper.MaxRowNumber &&
                       0 < ColumnNumber && ColumnNumber <= XLHelper.MaxColumnNumber;
            }
        }

        /// <summary>
        /// Get the sheet part of an address with a sheet. The address must have a worksheet.
        /// </summary>
        private string GetSheetPrefix()
        {
            if (Worksheet is null)
                throw new InvalidOperationException("The address has no worksheet.");

            return Worksheet.IsDeleted ? "#REF" : Worksheet.Name.EscapeSheetName();
        }
    }
}
