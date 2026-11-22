using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace ClosedXML.Excel
{
    internal readonly struct XLRangeAddress : IXLRangeAddress, IEquatable<XLRangeAddress>
    {
        #region Static members

        /// <summary>
        /// The space for the longest text of a range without a sheet name. The longest text is
        /// <c>#REF!$XFD$1048576:$XFD$1048576</c>.
        /// </summary>
        internal const int MaxTextLength = 5 + 2 * XLAddress.MaxA1Length + 1;

        private static readonly string EntireSheetText = $"1:{XLHelper.MaxRowNumber}";

        public static XLRangeAddress EntireColumn(XLWorksheet worksheet, int column)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, 1, column, false, false),
                new XLAddress(worksheet, XLHelper.MaxRowNumber, column, false, false));
        }

        public static XLRangeAddress EntireRow(XLWorksheet worksheet, int row)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, row, 1, false, false),
                new XLAddress(worksheet, row, XLHelper.MaxColumnNumber, false, false));
        }

        public static readonly XLRangeAddress Invalid = new XLRangeAddress(
            new XLAddress(-1, -1, fixedRow: true, fixedColumn: true),
            new XLAddress(-1, -1, fixedRow: true, fixedColumn: true)
        );

        internal static XLRangeAddress FromSheetRange(XLWorksheet? worksheet, Area range)
        {
            return new XLRangeAddress(
                new XLAddress(worksheet, range.FirstPoint.Row, range.FirstPoint.Column, fixedRow: false, fixedColumn: false),
                new XLAddress(range.LastPoint.Row, range.LastPoint.Column, fixedRow: false, fixedColumn: false)
            );
        }

        #endregion Static members

        #region Constructor

        public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress) : this()
        {
            Worksheet = firstAddress.Worksheet;
            FirstAddress = firstAddress;
            LastAddress = lastAddress;
        }

        public XLRangeAddress(XLWorksheet? worksheet, String rangeAddress) : this()
        {
            string addressToUse = rangeAddress.Contains('!')
                ? rangeAddress.Substring(rangeAddress.LastIndexOf('!') + 1)
                : rangeAddress;

            string firstPart;
            string secondPart;
            if (addressToUse.Contains(':'))
            {
                var arrRange = addressToUse.Split(':');
                firstPart = arrRange[0];
                secondPart = arrRange[1];
            }
            else
            {
                firstPart = addressToUse;
                secondPart = addressToUse;
            }

            if (XLHelper.IsValidA1Address(firstPart))
            {
                FirstAddress = XLAddress.Create(worksheet, firstPart);
                LastAddress = XLAddress.Create(worksheet, secondPart);
            }
            else
            {
                firstPart = firstPart.Replace("$", String.Empty);
                secondPart = secondPart.Replace("$", String.Empty);
                if (char.IsDigit(firstPart[0]))
                {
                    FirstAddress = XLAddress.Create(worksheet, "A" + firstPart);
                    LastAddress = XLAddress.Create(worksheet, XLHelper.MaxColumnLetter + secondPart);
                }
                else
                {
                    FirstAddress = XLAddress.Create(worksheet, firstPart + "1");
                    LastAddress = XLAddress.Create(worksheet, secondPart + XLHelper.MaxRowNumber.ToInvariantString());
                }
            }

            Worksheet = worksheet;
        }

        #endregion Constructor

        #region Public properties

        public XLWorksheet? Worksheet { get; }

        public XLAddress FirstAddress { get; }

        public XLAddress LastAddress { get; }

        IXLWorksheet? IXLRangeAddress.Worksheet
        {
            get { return Worksheet; }
        }

        IXLAddress IXLRangeAddress.FirstAddress
        {
            [DebuggerStepThrough]
            get { return FirstAddress; }
        }

        IXLAddress IXLRangeAddress.LastAddress
        {
            [DebuggerStepThrough]
            get { return LastAddress; }
        }

        public bool IsValid => FirstAddress.IsValid && LastAddress.IsValid;

        public int ColumnSpan
        {
            get
            {
                if (!IsValid)
                    throw new InvalidOperationException("Range address is invalid.");

                return Math.Abs(LastAddress.ColumnNumber - FirstAddress.ColumnNumber) + 1;
            }
        }

        public int NumberOfCells => ColumnSpan * RowSpan;

        internal long Size => ColumnSpan * (long)RowSpan;

        public int RowSpan
        {
            get
            {
                if (!IsValid)
                    throw new InvalidOperationException("Range address is invalid.");

                return Math.Abs(LastAddress.RowNumber - FirstAddress.RowNumber) + 1;
            }
        }

        private bool WorksheetIsDeleted => Worksheet?.IsDeleted == true;

        #endregion Public properties

        #region Public methods

        public Boolean IsNormalized => LastAddress.RowNumber >= FirstAddress.RowNumber
                                       && LastAddress.ColumnNumber >= FirstAddress.ColumnNumber;

        /// <summary>
        /// Lead a range address to a normal form - when <see cref="FirstAddress"/> points to the top-left address and
        /// <see cref="LastAddress"/> points to the bottom-right address.
        /// </summary>
        public XLRangeAddress Normalize()
        {
            if (FirstAddress.RowNumber <= LastAddress.RowNumber &&
                FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
                return this;

            int firstRow, firstColumn, lastRow, lastColumn;
            bool firstRowFixed, firstColumnFixed, lastRowFixed, lastColumnFixed;

            if (FirstAddress.RowNumber <= LastAddress.RowNumber)
            {
                firstRow = FirstAddress.RowNumber;
                firstRowFixed = FirstAddress.FixedRow;
                lastRow = LastAddress.RowNumber;
                lastRowFixed = LastAddress.FixedRow;
            }
            else
            {
                firstRow = LastAddress.RowNumber;
                firstRowFixed = LastAddress.FixedRow;
                lastRow = FirstAddress.RowNumber;
                lastRowFixed = FirstAddress.FixedRow;
            }

            if (FirstAddress.ColumnNumber <= LastAddress.ColumnNumber)
            {
                firstColumn = FirstAddress.ColumnNumber;
                firstColumnFixed = FirstAddress.FixedColumn;
                lastColumn = LastAddress.ColumnNumber;
                lastColumnFixed = LastAddress.FixedColumn;
            }
            else
            {
                firstColumn = LastAddress.ColumnNumber;
                firstColumnFixed = LastAddress.FixedColumn;
                lastColumn = FirstAddress.ColumnNumber;
                lastColumnFixed = FirstAddress.FixedColumn;
            }

            return new XLRangeAddress(
                new XLAddress(FirstAddress.Worksheet, firstRow, firstColumn, firstRowFixed, firstColumnFixed),
                new XLAddress(LastAddress.Worksheet, lastRow, lastColumn, lastRowFixed, lastColumnFixed));
        }

        public bool Intersects(IXLRangeAddress otherAddress)
        {
            var xlOtherAddress = (XLRangeAddress)otherAddress;
            return Intersects(in xlOtherAddress);
        }

        internal bool Intersects(in XLRangeAddress otherAddress)
        {
            return !( // See if the two ranges intersect...
                       otherAddress.FirstAddress.ColumnNumber > LastAddress.ColumnNumber
                    || otherAddress.LastAddress.ColumnNumber < FirstAddress.ColumnNumber
                    || otherAddress.FirstAddress.RowNumber > LastAddress.RowNumber
                    || otherAddress.LastAddress.RowNumber < FirstAddress.RowNumber
                );
        }

        public bool Contains(IXLAddress address)
        {
            var xlAddress = (XLAddress)address;
            return Contains(in xlAddress);
        }

        /// <summary>
        /// Does this range contains whole another range?
        /// </summary>
        public bool ContainsWhole(IXLRangeAddress range)
        {
            if (!range.IsValid)
                return false;

            return
                range.FirstAddress.ColumnNumber >= FirstAddress.ColumnNumber &&
                range.FirstAddress.RowNumber >= FirstAddress.RowNumber &&
                range.LastAddress.ColumnNumber <= LastAddress.ColumnNumber &&
                range.LastAddress.RowNumber <= LastAddress.RowNumber;
        }

        internal IXLRangeAddress WithoutWorksheet()
        {
            return new XLRangeAddress(
                FirstAddress.WithoutWorksheet(),
                LastAddress.WithoutWorksheet());
        }

        internal XLRangeAddress WithWorksheet(XLWorksheet worksheet)
        {
            return new XLRangeAddress(
                FirstAddress.WithWorksheet(worksheet),
                LastAddress.WithWorksheet(worksheet));
        }

        internal bool Contains(in XLAddress address)
        {
            return FirstAddress.RowNumber <= address.RowNumber &&
                   address.RowNumber <= LastAddress.RowNumber &&
                   FirstAddress.ColumnNumber <= address.ColumnNumber &&
                   address.ColumnNumber <= LastAddress.ColumnNumber;
        }

        public String ToStringRelative()
        {
            return ToStringRelative(false);
        }

        public String ToStringFixed()
        {
            return ToStringFixed(XLReferenceStyle.A1);
        }

        public String ToStringRelative(Boolean includeSheet)
        {
            Span<char> address = stackalloc char[MaxTextLength];
            var length = FormatRelative(address);

            if (includeSheet || WorksheetIsDeleted)
                return String.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet!.Name.EscapeSheetName(),
                    "!", address[..length]);

            return address[..length].ToString();
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle)
        {
            return ToStringFixed(referenceStyle, false);
        }

        public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
        {
            string address;
            if (!IsValid)
                address = "#REF!";
            else
            {
                if (IsEntireSheet())
                    address = $"$1:${XLHelper.MaxRowNumber}";
                else if (IsEntireRow())
                    address = String.Concat("$", FirstAddress.RowNumber.ToString(), ":$", LastAddress.RowNumber.ToString());
                else if (IsEntireColumn())
                    address = String.Concat("$", FirstAddress.ColumnLetter, ":$", LastAddress.ColumnLetter);
                else
                    address = String.Concat(FirstAddress.ToStringFixed(referenceStyle), ":", LastAddress.ToStringFixed(referenceStyle));
            }

            if (includeSheet || WorksheetIsDeleted)
                return String.Concat(
                    WorksheetIsDeleted ? "#REF" : Worksheet!.Name.EscapeSheetName(),
                    "!", address);

            return address;
        }

        public override string ToString()
        {
            Span<char> address = stackalloc char[MaxTextLength];
            var length = Format(address);
            return address[..length].ToString();
        }

        /// <summary>
        /// Write the text of <see cref="ToString()"/> to the span.
        /// </summary>
        /// <param name="output">The span. It must have space for <see cref="MaxTextLength"/> characters.</param>
        /// <returns>The number of characters that the method writes.</returns>
        internal int Format(Span<char> output)
        {
            if (!IsValid || WorksheetIsDeleted)
            {
                var length = WorksheetIsDeleted ? Write(output, XLHelper.RefError) : 0;
                return length + (IsValid ? FormatCells(output[length..]) : Write(output[length..], XLHelper.RefError));
            }

            return FormatValid(output, fixedReference: true);
        }

        /// <summary>
        /// Write the text of <see cref="ToStringRelative(bool)"/> without a sheet to the span.
        /// </summary>
        /// <param name="output">The span. It must have space for <see cref="MaxTextLength"/> characters.</param>
        /// <returns>The number of characters that the method writes.</returns>
        internal int FormatRelative(Span<char> output)
        {
            if (!IsValid)
                return Write(output, XLHelper.RefError);

            return FormatValid(output, fixedReference: false);
        }

        /// <summary>
        /// Write the first cell, a colon and the last cell to the span. A cell that is not valid
        /// is <c>#REF!</c>. The text has a <c>$</c> before a fixed row or column.
        /// </summary>
        /// <param name="output">The span. It must have space for <see cref="MaxTextLength"/> characters.</param>
        /// <returns>The number of characters that the method writes.</returns>
        internal int FormatCells(Span<char> output)
        {
            var length = FormatCell(output, FirstAddress, fixedReference: true);
            output[length++] = ':';
            return length + FormatCell(output[length..], LastAddress, fixedReference: true);
        }

        private int FormatValid(Span<char> output, bool fixedReference)
        {
            if (IsEntireSheet())
                return Write(output, EntireSheetText);

            int length;
            if (IsEntireRow())
            {
                var isFirstWritten = FirstAddress.RowNumber.TryFormat(output, out length, provider: CultureInfo.InvariantCulture);
                output[length++] = ':';
                var isLastWritten = LastAddress.RowNumber.TryFormat(output[length..], out var lastLength, provider: CultureInfo.InvariantCulture);
                Debug.Assert(isFirstWritten && isLastWritten, "The span is too short for the row numbers.");
                return length + lastLength;
            }

            if (IsEntireColumn())
            {
                length = Write(output, FirstAddress.ColumnLetter);
                output[length++] = ':';
                return length + Write(output[length..], LastAddress.ColumnLetter);
            }

            length = FormatCell(output, FirstAddress, fixedReference);
            output[length++] = ':';
            return length + FormatCell(output[length..], LastAddress, fixedReference);
        }

        private static int FormatCell(Span<char> output, XLAddress address, bool fixedReference)
        {
            if (!address.IsValid)
                return Write(output, XLHelper.RefError);

            return fixedReference
                ? address.FormatA1(output, address.FixedColumn, address.FixedRow)
                : address.FormatA1(output, fixedColumn: false, fixedRow: false);
        }

        private static int Write(Span<char> output, string text)
        {
            text.CopyTo(output);
            return text.Length;
        }

        public string ToString(XLReferenceStyle referenceStyle)
        {
            return ToString(referenceStyle, false);
        }

        public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
        {
            if (referenceStyle == XLReferenceStyle.R1C1)
                return ToStringFixed(referenceStyle, true);
            else
                return ToStringRelative(includeSheet);
        }

        public override bool Equals(object? obj)
        {
            if (!(obj is XLRangeAddress))
            {
                return false;
            }

            var address = (XLRangeAddress)obj;
            return FirstAddress.Equals(address.FirstAddress) &&
                   LastAddress.Equals(address.LastAddress) &&
                   EqualityComparer<XLWorksheet?>.Default.Equals(Worksheet, address.Worksheet);
        }

        public override int GetHashCode()
        {
            var hashCode = -778064135;
            hashCode = hashCode * -1521134295 + FirstAddress.GetHashCode();
            hashCode = hashCode * -1521134295 + LastAddress.GetHashCode();
            hashCode = hashCode * -1521134295 + (Worksheet?.GetHashCode() ?? 0);
            return hashCode;
        }

        public bool Equals(XLRangeAddress other)
        {
            return ReferenceEquals(Worksheet, other.Worksheet) &&
                   FirstAddress == other.FirstAddress &&
                   LastAddress == other.LastAddress;
        }

        public bool IsSingleCell()
        {
            return IsValid
                   && FirstAddress.RowNumber == LastAddress.RowNumber
                   && FirstAddress.ColumnNumber == LastAddress.ColumnNumber;
        }

        public bool IsEntireColumn()
        {
            return IsValid
                   && FirstAddress.RowNumber == 1
                   && LastAddress.RowNumber == XLHelper.MaxRowNumber;
        }

        public bool IsEntireRow()
        {
            return IsValid
                   && FirstAddress.ColumnNumber == 1
                   && LastAddress.ColumnNumber == XLHelper.MaxColumnNumber;
        }

        public bool IsEntireSheet()
        {
            return IsValid && IsEntireColumn() && IsEntireRow();
        }

        public IXLRangeAddress Relative(IXLRangeAddress sourceRangeAddress, IXLRangeAddress targetRangeAddress)
        {
            var xlSourceRangeAddress = (XLRangeAddress)sourceRangeAddress;
            var xlTargetRangeAddress = (XLRangeAddress)targetRangeAddress;

            return Relative(in xlSourceRangeAddress, in xlTargetRangeAddress);
        }

        internal XLRangeAddress Relative(in XLRangeAddress sourceRangeAddress, in XLRangeAddress targetRangeAddress)
        {
            var sheet = targetRangeAddress.Worksheet;

            return new XLRangeAddress
            (
                new XLAddress
                (
                    sheet,
                    this.FirstAddress.RowNumber - sourceRangeAddress.FirstAddress.RowNumber + targetRangeAddress.FirstAddress.RowNumber,
                    this.FirstAddress.ColumnNumber - sourceRangeAddress.FirstAddress.ColumnNumber + targetRangeAddress.FirstAddress.ColumnNumber,
                    fixedRow: false,
                    fixedColumn: false
                ),
                new XLAddress
                (
                    sheet,
                    this.LastAddress.RowNumber - sourceRangeAddress.FirstAddress.RowNumber + targetRangeAddress.FirstAddress.RowNumber,
                    this.LastAddress.ColumnNumber - sourceRangeAddress.FirstAddress.ColumnNumber + targetRangeAddress.FirstAddress.ColumnNumber,
                    fixedRow: false,
                    fixedColumn: false
                )
            );
        }

        public IXLRangeAddress Intersection(IXLRangeAddress otherRangeAddress)
        {
            if (otherRangeAddress == null)
                throw new ArgumentNullException(nameof(otherRangeAddress));

            var xlOtherRangeAddress = (XLRangeAddress)otherRangeAddress;
            return Intersection(in xlOtherRangeAddress);
        }

        internal XLRangeAddress Intersection(in XLRangeAddress otherRangeAddress)
        {
            if (!Equals(Worksheet, otherRangeAddress.Worksheet))
                throw new ArgumentOutOfRangeException(nameof(otherRangeAddress), "The other range address is on a different worksheet");

            var thisRangeAddressNormalized = this.Normalize();
            var otherRangeAddressNormalized = otherRangeAddress.Normalize();

            var firstRow = Math.Max(thisRangeAddressNormalized.FirstAddress.RowNumber, otherRangeAddressNormalized.FirstAddress.RowNumber);
            var firstColumn = Math.Max(thisRangeAddressNormalized.FirstAddress.ColumnNumber, otherRangeAddressNormalized.FirstAddress.ColumnNumber);
            var lastRow = Math.Min(thisRangeAddressNormalized.LastAddress.RowNumber, otherRangeAddressNormalized.LastAddress.RowNumber);
            var lastColumn = Math.Min(thisRangeAddressNormalized.LastAddress.ColumnNumber, otherRangeAddressNormalized.LastAddress.ColumnNumber);

            if (lastRow < firstRow || lastColumn < firstColumn)
                return XLRangeAddress.Invalid;

            return new XLRangeAddress
            (
                new XLAddress(this.Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
                new XLAddress(this.Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
            );
        }

        public IXLRange? AsRange()
        {
            if (this.Worksheet == null)
                throw new InvalidOperationException("The worksheet of the current range address has not been set.");

            if (!this.IsValid)
                return null;

            return this.Worksheet.Range(this);
        }

        #endregion Public methods

        #region Operators

        public static bool operator ==(XLRangeAddress left, XLRangeAddress right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(XLRangeAddress left, XLRangeAddress right)
        {
            return !(left == right);
        }

        #endregion Operators
    }
}
