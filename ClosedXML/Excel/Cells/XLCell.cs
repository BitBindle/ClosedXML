using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Visitors;
using ClosedXML.Excel.Formatting;
using ClosedXML.Excel.InsertData;
using ClosedXML.Extensions;
using ClosedXML.Graphics;
using ClosedXML.Parser;

namespace ClosedXML.Excel
{
    [DebuggerDisplay("{Address}")]
    internal sealed class XLCell : IXLCell, IXLFormatContainer
    {
        private readonly XLCellsCollection _cellsCollection;

        private readonly int _rowNumber;

        private readonly int _columnNumber;

        internal XLCell(XLWorksheet worksheet, int row, int column)
        {
            _cellsCollection = worksheet.Internals.CellsCollection;
            _rowNumber = row;
            _columnNumber = column;
        }

        internal XLCell(XLWorksheet worksheet, Point point) : this(worksheet, point.Row, point.Column)
        {
        }

        public XLWorksheet Worksheet => _cellsCollection.Worksheet;

        public XLAddress Address => new(Worksheet, _rowNumber, _columnNumber, false, false);

        internal Point Point => new(_rowNumber, _columnNumber);

        private XLWorkbookStyles Styles => Worksheet.Workbook.Styles;

        #region Slice fields

        /// <summary>
        /// A flag indicating if a string should be stored in the shared table or inline.
        /// </summary>
        public bool ShareString
        {
            get => _cellsCollection.ValueSlice.GetShareString(Point);
            set => _cellsCollection.ValueSlice.SetShareString(Point, value);
        }

        internal int MemorySstId => _cellsCollection.ValueSlice.GetShareStringId(Point);

        internal XLImmutableRichText? RichText => SliceRichText;

        private XLCellValue SliceCellValue
        {
            get => _cellsCollection.ValueSlice.GetCellValue(Point);
            set
            {
                _cellsCollection.ValueSlice.SetCellValue(Point, value);
                Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, Point);
            }
        }

        [DisallowNull]
        private XLImmutableRichText? SliceRichText
        {
            get => _cellsCollection.ValueSlice.GetRichText(Point);
            set => _cellsCollection.ValueSlice.SetRichText(Point, value);
        }

        private XLComment? SliceComment
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].Comment;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.Comment != value)
                {
                    var modified = original;
                    modified.Comment = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        internal UInt32? CellMetaIndex
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].CellMetaIndex;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.CellMetaIndex != value)
                {
                    var modified = original;
                    modified.CellMetaIndex = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        internal UInt32? ValueMetaIndex
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].ValueMetaIndex;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.ValueMetaIndex != value)
                {
                    var modified = original;
                    modified.ValueMetaIndex = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        /// <summary>
        /// A formula in the cell. Null, if cell doesn't contain formula.
        /// </summary>
        internal XLCellFormula? Formula
        {
            get => _cellsCollection.FormulaSlice.Get(Point);
            set
            {
                _cellsCollection.FormulaSlice.Set(Point, value);

                // Because text values of evaluated formulas are stored in a worksheet part, mark it as inlined string and store in sst.
                // If we are clearing formula, we should enable shareString back on, because it is a default position.
                // If we are setting formula, we should disable shareString (=inline), because it must be written to the worksheet part
                var clearFormula = value is null;
                ShareString = clearFormula;
                Worksheet.Workbook.CalcEngine.MarkDirty(Worksheet, Point);
            }
        }

        #endregion Slice fields

        #region IXLFormatContainer

        /// <summary>
        /// A format of a cell. If cell format depends on inherited format, the value is <c>null</c>.
        /// </summary>
        public XLCellFormatValue? FormatValue
        {
            get => _cellsCollection.FormatSlice.GetFormat(Point);
            set => _cellsCollection.FormatSlice.Set(Point, value);
        }
        #endregion

        internal XLCellFormat Format => XLCellFormat.ForCell(this);

        internal XLComment GetComment()
        {
            return SliceComment ?? CreateComment();
        }

        internal XLComment CreateComment(int? shapeId = null)
        {
            return SliceComment = XLComment.Create(this, shapeId: shapeId);
        }

        public XLRichText GetRichText()
        {
            var sliceRichText = SliceRichText;
            if (sliceRichText is not null)
            {
                var cellFormat = GetFormat();
                return new XLRichText(this, cellFormat.Font, sliceRichText);
            }

            return CreateRichText();
        }

        public XLRichText CreateRichText()
        {
            var fontFormat = GetFormat().Font;

            // Don't include rich text string with 0 length to a new rich text
            var richText = DataType == XLDataType.Blank
                ? new XLRichText(this, fontFormat)
                : new XLRichText(this, fontFormat, GetFormattedString());
            SliceRichText = XLImmutableRichText.Create(richText);
            return richText;
        }

        #region IXLCell Members

        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        IXLWorksheet IXLCell.Worksheet
        {
            get { return Worksheet; }
        }

        IXLAddress IXLCell.Address
        {
            get { return Address; }
        }

        IXLRange IXLCell.AsRange()
        {
            return AsRange();
        }

        internal IXLCell SetValue(XLCellValue value, bool setTableHeader, bool checkMergedRanges)
        {
            if (checkMergedRanges && IsInferiorMergedCell())
                return this;

            // Mimic Excel behavior: When a value is set to a a certain types (e.g. timespan or
            // a date), the format of a cell is changed.
            var valueRequiredFormat = Worksheet.GetStyleForValue(value, Point);
            if (valueRequiredFormat is not null)
                FormatValue = valueRequiredFormat;

            // Modify value after style, because we might need to strip the '
            if (value.Type == XLDataType.Text)
            {
                var text = value.GetText();
                if (text.Length > 0 && text[0] == '\'')
                {
                    value = text.Substring(1);
                }
            }

            SetOnlyValue(value);

            FormulaA1 = null;

            if (setTableHeader)
            {
                var cellRange = new Area(Point, Point);
                foreach (var table in Worksheet.Tables)
                    table.RefreshFieldsFromCells(cellRange);
            }

            return this;
        }

        public Boolean GetBoolean() => Value.GetBoolean();

        public Double GetDouble() => Value.GetNumber();

        public string GetText() => Value.GetText();

        public XLError GetError() => Value.GetError();

        public DateTime GetDateTime() => Value.GetDateTime();

        public TimeSpan GetTimeSpan() => Value.GetTimeSpan();

        /// <summary>
        /// 2^63, the smallest double that is more than <see cref="Int64.MaxValue"/>.
        /// </summary>
        private const double Int64UpperLimit = 9223372036854775808d;

        /// <summary>
        /// 2^64, the smallest double that is more than <see cref="UInt64.MaxValue"/>.
        /// </summary>
        private const double UInt64UpperLimit = 18446744073709551616d;

        public Boolean TryGetValue<T>([MaybeNullWhen(false)] out T value)
        {
            XLCellValue currentValue;
            try
            {
                currentValue = Value;
            }
            catch
            {
                // May fail for formula evaluation
                value = default;
                return false;
            }

            var targetType = typeof(T);
            var isNullable = targetType.IsNullableType();
            if (isNullable && currentValue.TryConvert(out Blank? _))
            {
                // T is a Nullable<> type here, thus null is a valid value of T.
                value = default!;
                return true;
            }

            // JIT compiles a separate version for each T value type and one for all reference types
            // Optimization then removes the double casting for value types.
            var underlyingType = targetType.GetUnderlyingType();
            if (underlyingType == typeof(DateTime) && currentValue.TryConvert(out DateTime dateTime))
            {
                value = (T)(object)dateTime;
                return true;
            }

            var culture = CultureInfo.CurrentCulture;
            if (underlyingType == typeof(TimeSpan) && currentValue.TryConvert(out TimeSpan timeSpan, culture))
            {
                value = (T)(object)timeSpan;
                return true;
            }

            if (underlyingType == typeof(Boolean) && currentValue.TryConvert(out Boolean boolean))
            {
                value = (T)(object)boolean;
                return true;
            }

            if (TryGetStringValue(out value, currentValue)) return true;

            if (underlyingType == typeof(XLError))
            {
                if (currentValue.IsError)
                {
                    value = (T)(object)currentValue.GetError();
                    return true;
                }

                return false;
            }

            // Type code of an enum is a type of an integer, so do this check before numbers
            if (underlyingType.IsEnum)
            {
                var strValue = currentValue.ToString(culture);
                if (Enum.IsDefined(underlyingType, strValue))
                {
                    value = (T)Enum.Parse(underlyingType, strValue, ignoreCase: false);
                    return true;
                }
                value = default;
                return false;
            }

            var typeCode = Type.GetTypeCode(underlyingType);

            // T is a floating point numbers
            if (typeCode >= TypeCode.Single && typeCode <= TypeCode.Decimal)
            {
                if (!currentValue.TryConvert(out Double doubleValue, culture))
                    return false;

                if (typeCode == TypeCode.Single && doubleValue is < Single.MinValue or > Single.MaxValue)
                    return false;

                if (typeCode == TypeCode.Decimal)
                {
                    // On .NET, the cast throws for each double of 2^96 or more. The limit on .NET
                    // Framework is not measured, so catch the overflow instead of a constant check.
                    try
                    {
                        value = (T)(object)(Decimal)doubleValue;
                        return true;
                    }
                    catch (OverflowException)
                    {
                        return false;
                    }
                }

                value = typeCode switch
                {
                    TypeCode.Single => (T)(object)(Single)doubleValue,
                    TypeCode.Double => (T)(object)doubleValue,
                    _ => throw new NotSupportedException()
                };
                return true;
            }

            // T is an integer
            if (typeCode >= TypeCode.SByte && typeCode <= TypeCode.UInt64)
            {
                if (!currentValue.TryConvert(out Double doubleValue, culture))
                    return false;

                if (!doubleValue.Equals(Math.Truncate(doubleValue)))
                    return false;

                var valueIsWithinBounds = typeCode switch
                {
                    TypeCode.SByte => doubleValue >= SByte.MinValue && doubleValue <= SByte.MaxValue,
                    TypeCode.Byte => doubleValue >= Byte.MinValue && doubleValue <= Byte.MaxValue,
                    TypeCode.Int16 => doubleValue >= Int16.MinValue && doubleValue <= Int16.MaxValue,
                    TypeCode.UInt16 => doubleValue >= UInt16.MinValue && doubleValue <= UInt16.MaxValue,
                    TypeCode.Int32 => doubleValue >= Int32.MinValue && doubleValue <= Int32.MaxValue,
                    TypeCode.UInt32 => doubleValue >= UInt32.MinValue && doubleValue <= UInt32.MaxValue,
                    // Int64.MaxValue and UInt64.MaxValue become 2^63 and 2^64 as a double, which
                    // are outside the range. The upper limit must exclude them.
                    TypeCode.Int64 => doubleValue >= Int64.MinValue && doubleValue < Int64UpperLimit,
                    TypeCode.UInt64 => doubleValue >= UInt64.MinValue && doubleValue < UInt64UpperLimit,
                    _ => throw new NotSupportedException()
                };
                if (!valueIsWithinBounds)
                    return false;

                value = typeCode switch
                {
                    TypeCode.SByte => (T)(object)(SByte)doubleValue,
                    TypeCode.Byte => (T)(object)(Byte)doubleValue,
                    TypeCode.Int16 => (T)(object)(Int16)doubleValue,
                    TypeCode.UInt16 => (T)(object)(UInt16)doubleValue,
                    TypeCode.Int32 => (T)(object)(Int32)doubleValue,
                    TypeCode.UInt32 => (T)(object)(UInt32)doubleValue,
                    TypeCode.Int64 => (T)(object)(Int64)doubleValue,
                    TypeCode.UInt64 => (T)(object)(UInt64)doubleValue,
                    _ => throw new NotSupportedException()
                };
                return true;
            }

            return false;
        }

        private static bool TryGetStringValue<T>([MaybeNullWhen(false)] out T value, XLCellValue currentValue)
        {
            if (typeof(T) == typeof(String))
            {
                // The cell value holds the real text. A load decodes the escaped characters, for
                // example _x000D_, before it sets the value.
                value = (T)(object)currentValue.ToString(CultureInfo.CurrentCulture);
                return true;
            }
            value = default;
            return false;
        }

        public T GetValue<T>()
        {
            if (TryGetValue<T>(out var retVal))
                return retVal;

            throw new InvalidCastException($"Cannot convert {Address.ToStringRelative(true)}'s value to " + typeof(T));
        }

        public String GetString() => Value.ToString(CultureInfo.CurrentCulture);

        public string GetFormattedString(CultureInfo? culture = null)
        {
            XLCellValue value;
            try
            {
                // Need to get actual value because formula might be out of date or value wasn't set at all
                // Unimplemented functions and features throw exceptions
                value = Value;
            }
            catch
            {
                value = CachedValue;
            }

            return GetFormattedString(value, culture);
        }

        internal string GetFormattedString(XLCellValue value, CultureInfo? culture = null)
        {
            culture ??= CultureInfo.CurrentCulture;
            var numberFormat = GetFormat().NumberFormat;
            return value.IsUnifiedNumber
                ? value.GetUnifiedNumber().ToExcelFormat(numberFormat, culture)
                : value.ToString(culture);
        }

        public void InvalidateFormula()
        {
            if (Formula is null)
            {
                return;
            }

            Formula.IsDirty = true;
        }

        /// <summary>
        /// Perform an evaluation of cell formula. If cell does not contain formula nothing happens, if cell does not need
        /// recalculation (<see cref="NeedsRecalculation"/> is False) nothing happens either, unless <paramref name="force"/> flag is specified.
        /// Otherwise recalculation is performed, result value is preserved in <see cref="CachedValue"/> and returned.
        /// </summary>
        /// <param name="force">Flag indicating whether a recalculation must be performed even is cell does not need it.</param>
        /// <returns>Null if cell does not contain a formula. Calculated value otherwise.</returns>
        public void Evaluate(Boolean force)
        {
            if (Formula is null)
            {
                return;
            }

            var shouldRecalculate = force || NeedsRecalculation;
            if (!shouldRecalculate)
            {
                return;
            }

            // TODO: Only one cell, somehow
            var wb = Worksheet.Workbook;
            wb.CalcEngine.Recalculate(wb, null);
        }

        /// <summary>
        /// Set only value, don't clear formula, don't set format.
        /// Sets the value even for merged cells.
        /// </summary>
        internal void SetOnlyValue(XLCellValue value)
        {
            SliceCellValue = value;
        }

        public IXLCell SetValue(XLCellValue value)
        {
            return SetValue(value, true, true);
        }

        public override string ToString() => ToString("A");

        public string ToString(string format)
        {
            return (format.ToUpper()) switch
            {
                "A" => this.Address.ToString(),
                "F" => HasFormula ? this.FormulaA1 : string.Empty,
                "NF" => Style.NumberFormat.Format,
                "FG" => Style.Font.FontColor.ToString(),
                "BG" => Style.Fill.BackgroundColor.ToString(),
                "V" => GetFormattedString(),
                _ => throw new FormatException($"Format {format} was not recognised."),
            };
        }

        public XLCellValue Value
        {
            get
            {
                if (Formula is not null)
                {
                    Evaluate(false);
                }

                return SliceCellValue;
            }
            set => SetValue(value);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data)
        {
            return InsertTable(data, null, true);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, bool createTable)
        {
            return InsertTable(data, null, createTable);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, String? tableName)
        {
            return InsertTable(data, tableName, true);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, String? tableName, Boolean createTable)
        {
            return InsertTable(data, tableName, createTable, addHeadings: true, transpose: false);
        }

        public IXLTable InsertTable<T>(IEnumerable<T> data, String? tableName, Boolean createTable, Boolean addHeadings, Boolean transpose)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return Worksheet.InsertTable(Point, reader, tableName, createTable, addHeadings, transpose);
        }

        public IXLTable? InsertTable(DataTable? data)
        {
            return InsertTable(data, null, true);
        }

        public IXLTable? InsertTable(DataTable? data, Boolean createTable)
        {
            return InsertTable(data, null, createTable);
        }

        public IXLTable? InsertTable(DataTable? data, String? tableName)
        {
            return InsertTable(data, tableName, true);
        }

        public IXLTable? InsertTable(DataTable? data, String? tableName, Boolean createTable)
        {
            if (data == null || data.Columns.Count == 0)
                return null;

            if (XLHelper.IsValidA1Address(tableName) || XLHelper.IsValidRCAddress(tableName))
                throw new InvalidOperationException($"Table name cannot be a valid Cell Address '{tableName}'.");

            if (createTable && this.Worksheet.Tables.Any<XLTable>(t => t.Contains(this)))
                throw new InvalidOperationException($"This cell '{this.Address}' is already part of a table.");

            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return Worksheet.InsertTable(Point, reader, tableName, createTable, addHeadings: true, transpose: false);
        }

        public XLTableCellType TableCellType()
        {
            var table = this.Worksheet.Tables.FirstOrDefault<XLTable>(t => t.AsRange().Contains(this));
            if (table == null) return XLTableCellType.None;

            if (table.HeadersRow() is { } headersRow && headersRow.RowNumber() == _rowNumber) return XLTableCellType.Header;
            if (table.TotalsRow() is { } totalsRow && totalsRow.RowNumber() == _rowNumber) return XLTableCellType.Total;

            return XLTableCellType.Data;
        }

        public IXLRange? InsertData(IEnumerable? data)
        {
            return InsertData(data, transpose: false);
        }

        public IXLRange? InsertData(IEnumerable? data, Boolean transpose)
        {
            if (data == null || data is String)
                return null;

            return InsertDataInternal(data, transpose);
        }

        /// <summary>
        /// Inserts the data elements and returns the range it occupies. The caller checks that the data is not a string.
        /// </summary>
        internal XLRange InsertDataInternal(IEnumerable data, Boolean transpose)
        {
            var reader = InsertDataReaderFactory.Instance.CreateReader(data);
            return Worksheet.InsertData(Point, reader, addHeadings: false, transpose: transpose);
        }

        public IXLRange? InsertData(DataTable? dataTable)
        {
            if (dataTable == null)
                return null;

            var reader = InsertDataReaderFactory.Instance.CreateReader(dataTable);
            return Worksheet.InsertData(Point, reader, addHeadings: false, transpose: false);
        }

        public XLDataType DataType => SliceCellValue.Type;

        public IXLCell Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            return Clear(clearOptions, false);
        }

        internal IXLCell Clear(XLClearOptions clearOptions, bool calledFromRange)
        {
            //Note: We have to check if the cell is part of a merged range. If so we have to clear the whole range
            //Checking if called from range to avoid stack overflow
            if (!calledFromRange && IsMerged())
            {
                var firstOrDefault = Worksheet.Internals.MergedRanges.GetIntersectedRanges(Address).FirstOrDefault();
                if (firstOrDefault != null)
                    firstOrDefault.Clear(clearOptions);
            }
            else
            {
                if (!calledFromRange)
                    ArrayFormulaShiftCheck.ThrowIfClearSplits(this, clearOptions);

                if (clearOptions.HasFlag(XLClearOptions.Contents))
                {
                    SetHyperlink(null);
                    SliceCellValue = Blank.Value;
                    FormulaA1 = String.Empty;
                }

                if (clearOptions.HasFlag(XLClearOptions.NormalFormats))
                    FormatValue = Worksheet.FormatValue;

                if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
                {
                    Worksheet.ConditionalFormats.Clear(Point);
                }

                if (clearOptions.HasFlag(XLClearOptions.Comments))
                    SliceComment = null;

                if (clearOptions.HasFlag(XLClearOptions.Sparklines))
                {
                    AsRange().RemoveSparklines();
                }

                if (clearOptions.HasFlag(XLClearOptions.DataValidation) && HasDataValidation)
                {
                    var validation = CreateDataValidation();
                    Worksheet.DataValidations.Delete(validation);
                }

                if (clearOptions.HasFlag(XLClearOptions.MergedRanges) && IsMerged())
                {
                    ClearMerged();
                }
            }

            return this;
        }

        public void Delete(XLShiftDeletedCells shiftDeleteCells)
        {
            Worksheet.Range(Address, Address).Delete(shiftDeleteCells);
        }

        [AllowNull]
        public string FormulaA1
        {
            get => Formula?.A1 ?? String.Empty;

            set
            {
                if (IsInferiorMergedCell())
                    return;

                var formula = value?.TrimFormulaEqual();
                if (!String.IsNullOrWhiteSpace(formula))
                {
                    var fixedFunctionsFormula = FormulaTransformation.FixFutureFunctions(formula, Worksheet.Name, Point);
                    Formula = XLCellFormula.NormalA1(fixedFunctionsFormula);
                }
                else
                {
                    Formula = null;
                }

                InvalidateFormula();
            }
        }

        [AllowNull]
        public string FormulaR1C1
        {
            get => Formula?.GetFormulaR1C1(Point) ?? String.Empty;

            set
            {
                if (IsInferiorMergedCell())
                    return;

                var formula = value?.TrimFormulaEqual();
                if (!String.IsNullOrWhiteSpace(formula))
                {
                    var formulaA1 = FormulaConverter.ToA1(formula, _rowNumber, _columnNumber);
                    var fixedFunctionsFormulaA1 = FormulaTransformation.FixFutureFunctions(formulaA1, Worksheet.Name, Point);
                    Formula = XLCellFormula.NormalA1(fixedFunctionsFormulaA1);
                }
                else
                {
                    Formula = null;
                }

                InvalidateFormula();
            }
        }

        public XLHyperlink GetHyperlink()
        {
            if (Worksheet.Hyperlinks.TryGet(Point, out var hyperlink))
                return hyperlink;

            return CreateHyperlink();
        }

        /// <inheritdoc />
        public void SetHyperlink(XLHyperlink? hyperlink)
        {
            Worksheet.Hyperlinks.SetCellHyperlink(Point, hyperlink);
            if (hyperlink is null)
                return;

            var cellFormat = GetFormat();
            var sheetFormat = Worksheet.GetFormat();
            if (ReferenceEquals(cellFormat, sheetFormat))
            {
                FormatValue = Styles.GetModifiedFormat(cellFormat, font => font with
                {
                    Color = XLColor.FromTheme(XLThemeColor.Hyperlink),
                    Underline = XLFontUnderlineValues.Single
                });
            }
        }

        internal void SetCellHyperlink(XLHyperlink hyperlink)
        {
            Worksheet.Hyperlinks.SetCellHyperlink(Point, hyperlink);
        }

        public XLHyperlink CreateHyperlink()
        {
            var link = new XLHyperlink();
            SetHyperlink(link);
            return link;
        }

        public IXLCells InsertCellsAbove(int numberOfRows)
        {
            return AsRange().InsertRowsAbove(numberOfRows).Cells();
        }

        public IXLCells InsertCellsBelow(int numberOfRows)
        {
            return AsRange().InsertRowsBelow(numberOfRows).Cells();
        }

        public IXLCells InsertCellsAfter(int numberOfColumns)
        {
            return AsRange().InsertColumnsAfter(numberOfColumns).Cells();
        }

        public IXLCells InsertCellsBefore(int numberOfColumns)
        {
            return AsRange().InsertColumnsBefore(numberOfColumns).Cells();
        }

        public IXLCell AddToNamed(string rangeName)
        {
            AsRange().AddToNamed(rangeName);
            return this;
        }

        public IXLCell AddToNamed(string rangeName, XLScope scope)
        {
            AsRange().AddToNamed(rangeName, scope);
            return this;
        }

        public IXLCell AddToNamed(string rangeName, XLScope scope, string comment)
        {
            AsRange().AddToNamed(rangeName, scope, comment);
            return this;
        }

        /// <summary>
        /// Flag indicating that previously calculated cell value may be not valid anymore and has to be re-evaluated.
        /// </summary>
        public bool NeedsRecalculation => Formula is not null && Formula.IsDirty;

        public XLCellValue CachedValue => SliceCellValue;

        IXLRichText IXLCell.GetRichText() => GetRichText();

        public bool HasRichText => SliceRichText is not null;

        IXLRichText IXLCell.CreateRichText() => CreateRichText();

        IXLComment IXLCell.GetComment() => GetComment();

        public bool HasComment
        {
            get { return SliceComment != null; }
        }

        IXLComment IXLCell.CreateComment()
        {
            return CreateComment(shapeId: null);
        }

        public Boolean IsMerged()
        {
            return Worksheet.Internals.MergedRanges.Contains(this);
        }

        public IXLRange? MergedRange()
        {
            return Worksheet
                .Internals
                .MergedRanges
                .GetIntersectedRanges(this)
                .FirstOrDefault();
        }

        public Boolean IsEmpty()
        {
            return IsEmpty(XLCellsUsedOptions.AllContents);
        }

        public Boolean IsEmpty(XLCellsUsedOptions options)
        {
            var isValueEmpty = SliceCellValue.Type switch
            {
                XLDataType.Blank => true,
                XLDataType.Text => SliceCellValue.GetText().Length == 0,
                _ => false
            };

            if (!isValueEmpty || HasFormula)
                return false;

            if (options.HasFlag(XLCellsUsedOptions.NormalFormats))
            {
                if (FormatValue is { } cellFormat)
                {
                    var inheritedFormat = Worksheet.GetInheritedFormat(Point);
                    if (!XLCellFormatValue.AreSame(cellFormat, inheritedFormat))
                        return false;
                }
            }

            if (options.HasFlag(XLCellsUsedOptions.MergedRanges) && IsMerged())
                return false;

            if (options.HasFlag(XLCellsUsedOptions.Comments) && HasComment)
                return false;

            if (options.HasFlag(XLCellsUsedOptions.DataValidation) && HasDataValidation)
                return false;

            if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats)
                && Worksheet.ConditionalFormats.SelectMany<XLConditionalFormat, IXLRange>(cf => cf.Ranges).Any(range => range.Contains(this)))
                return false;

            if (options.HasFlag(XLCellsUsedOptions.Sparklines) && HasSparkline)
                return false;

            return true;
        }

        public IXLColumn WorksheetColumn()
        {
            return Worksheet.Column(_columnNumber);
        }

        public IXLRow WorksheetRow()
        {
            return Worksheet.Row(_rowNumber);
        }

        public IXLCell CopyTo(IXLCell target)
        {
            ((XLCell)target).CopyFrom(this, XLCellCopyOptions.All);
            return target;
        }

        public IXLCell CopyTo(String target)
        {
            return CopyTo(GetTargetCell(target, Worksheet));
        }

        public IXLCell CopyFrom(IXLCell otherCell)
        {
            return CopyFrom(otherCell, XLCellCopyOptions.All);
        }

        public IXLCell CopyFrom(String otherCell)
        {
            return CopyFrom(GetTargetCell(otherCell, Worksheet));
        }

        public IXLCell SetFormulaA1(String formula)
        {
            FormulaA1 = formula;
            return this;
        }

        public IXLCell SetFormulaR1C1(String formula)
        {
            FormulaR1C1 = formula;
            return this;
        }

        public Boolean HasSparkline => Sparkline != null;

        /// <summary> The sparkline assigned to the cell </summary>
        public IXLSparkline? Sparkline => Worksheet.SparklineGroups.GetSparkline(this);

        public IXLDataValidation GetDataValidation()
        {
            return FindDataValidation() ?? CreateDataValidation();
        }

        public Boolean HasDataValidation
        {
            get { return FindDataValidation() != null; }
        }

        /// <summary>
        /// Get the data validation rule containing current cell.
        /// </summary>
        /// <returns>The data validation rule applying to the current cell or null if there is no such rule.</returns>
        private IXLDataValidation? FindDataValidation()
        {
            Worksheet.DataValidations.TryGet(new XLRangeAddress(Address, Address), out var dataValidation);
            return dataValidation;
        }

        IXLDataValidation IXLCell.CreateDataValidation() => CreateDataValidation();

        internal XLDataValidation CreateDataValidation()
        {
            return Worksheet.DataValidations.Create(new Area(Point));
        }

        public void Select()
        {
            AsRange().Select();
        }

        public IXLConditionalFormat AddConditionalFormat()
        {
            return AsRange().AddConditionalFormat();
        }

        public Boolean Active
        {
            get => Worksheet.ActiveCell == Point;
            set
            {
                if (value)
                    Worksheet.ActiveCell = Point;
                else if (Active)
                    Worksheet.ActiveCell = null;
            }
        }

        public IXLCell SetActive(Boolean value = true)
        {
            Active = value;
            return this;
        }

        public Boolean HasHyperlink => Worksheet.Hyperlinks.HasHyperlink(Point);

        /// <inheritdoc />
        public Boolean ShowPhonetic
        {
            get => _cellsCollection.MiscSlice[_rowNumber, _columnNumber].HasPhonetic;
            set
            {
                ref readonly var original = ref _cellsCollection.MiscSlice[_rowNumber, _columnNumber];
                if (original.HasPhonetic != value)
                {
                    var modified = original;
                    modified.HasPhonetic = value;
                    _cellsCollection.MiscSlice.Set(_rowNumber, _columnNumber, in modified);
                }
            }
        }

        #endregion IXLCell Members

        /// <summary>
        /// Ensure the cell has style set directly on the cell, not inherited from column/row/worksheet styles.
        /// </summary>
        internal void PingStyle()
        {
            FormatValue = GetFormat();
        }

        public XLRange AsRange()
        {
            return Worksheet.Range(Address, Address);
        }

        #region Styles

        /// <summary>
        /// Get format of a cell that should be used to render it.
        /// </summary>
        internal XLCellFormatValue GetFormat()
        {
            return Worksheet.GetStyleValue(Point);
        }

        #endregion Styles

        public void DeleteComment()
        {
            Clear(XLClearOptions.Comments);
        }

        public void DeleteSparkline()
        {
            Clear(XLClearOptions.Sparklines);
        }

        public IXLCell CopyFrom(IXLRangeBase rangeObject)
        {
            if (rangeObject is null)
                throw new ArgumentNullException(nameof(rangeObject));

            var asRange = (XLRangeBase)rangeObject;
            return CopyFrom(asRange, PrepareRangeCopy(asRange));
        }

        /// <summary>
        /// Check a copy of a range to this cell and create the copies of its array formulas. Call it before
        /// the first change to the target, because it throws when the copy is not possible. The target can
        /// also overlap the source.
        /// </summary>
        /// <param name="source">The copied range.</param>
        /// <returns>The array formulas for <see cref="CopyFrom(XLRangeBase, List{XLCellFormula})"/>.</returns>
        internal List<XLCellFormula> PrepareRangeCopy(XLRangeBase source)
        {
            ThrowIfCopyIsOutsideSheet(source);
            return ArrayFormulaCopy.CreateCopies(source.Worksheet, source.SheetRange, Worksheet, Point);
        }

        /// <summary>
        /// Copy a range to this cell.
        /// </summary>
        /// <param name="asRange">The copied range.</param>
        /// <param name="arrayFormulaCopies">The array formulas from <see cref="PrepareRangeCopy"/>.</param>
        internal IXLCell CopyFrom(XLRangeBase asRange, List<XLCellFormula> arrayFormulaCopies)
        {
            var maxRows = asRange.RowCount();
            var maxColumns = asRange.ColumnCount();

            var lastRow = Math.Min(_rowNumber + maxRows - 1, XLHelper.MaxRowNumber);
            var lastColumn = Math.Min(_columnNumber + maxColumns - 1, XLHelper.MaxColumnNumber);

            var targetRange = Worksheet.Range(_rowNumber, _columnNumber, lastRow, lastColumn);

            if (!(asRange is XLRow || asRange is XLColumn))
            {
                targetRange.Clear();
            }

            var minRow = asRange.RangeAddress.FirstAddress.RowNumber;
            var minColumn = asRange.RangeAddress.FirstAddress.ColumnNumber;
            var cellsUsed = asRange.CellsUsed(CopiedCellsUsedOptions);
            foreach (var sourceCell in cellsUsed)
            {
                Worksheet.Cell(
                    _rowNumber + sourceCell.Address.RowNumber - minRow,
                    _columnNumber + sourceCell.Address.ColumnNumber - minColumn
                    ).CopyFromInternal((XLCell)sourceCell,
                    XLCellCopyOptions.All
                        & ~XLCellCopyOptions.ConditionalFormats
                        & ~XLCellCopyOptions.DataValidations); //Conditional formats and data validation are copied separately
            }

            // The cell copy copies only the values of an array formula. A data table also stays as values, like in Excel.
            ArrayFormulaCopy.SetArrayFormulas(Worksheet, arrayFormulaCopies);

            var rangesToMerge = asRange.Worksheet.Internals.MergedRanges
                .Where<XLRange>(mr => asRange.Contains(mr))
                .Select(mr =>
                {
                    var firstRow = _rowNumber + (mr.RangeAddress.FirstAddress.RowNumber - asRange.RangeAddress.FirstAddress.RowNumber);
                    var firstColumn = _columnNumber + (mr.RangeAddress.FirstAddress.ColumnNumber - asRange.RangeAddress.FirstAddress.ColumnNumber);
                    return (IXLRange)Worksheet.Range
                    (
                        firstRow,
                        firstColumn,
                        firstRow + mr.RowCount() - 1,
                        firstColumn + mr.ColumnCount() - 1
                    );
                })
                .ToList();

            rangesToMerge.ForEach(r => r.Merge(false));

            var dataValidations = asRange.Worksheet.DataValidations
                .GetAllInRange(asRange.RangeAddress)
                .ToList();

            foreach (var dataValidation in dataValidations)
            {
                XLDataValidation? newDataValidation = null;
                foreach (var dvRange in dataValidation.Ranges.Where(r => r.Intersects(asRange)))
                {
                    // Copy only the part of the validation in the source. The copy of an empty cell can be after
                    // the last row or the last column, so remove that part of the validation.
                    var copiedArea = Area.FromRangeAddress(dvRange.RangeAddress).Intersect(asRange.SheetRange);
                    if (copiedArea?.ShiftAndClip(_rowNumber - minRow, _columnNumber - minColumn) is not { } dvTargetArea)
                        continue;

                    var dvTargetRange = Worksheet.Range(dvTargetArea.TopRow, dvTargetArea.LeftColumn, dvTargetArea.BottomRow, dvTargetArea.RightColumn);
                    if (newDataValidation == null)
                    {
                        newDataValidation = dvTargetRange.CreateDataValidation();
                        newDataValidation.CopyFrom(dataValidation);
                    }
                    else
                        newDataValidation.AddRange(dvTargetRange);
                }
            }

            Worksheet.ConditionalFormats.CopyFrom(asRange.Worksheet, asRange.SheetRange, Point);
            return this;
        }

        /// <summary>
        /// The options that select the source cells of a range copy. The copy copies conditional formats,
        /// data validations and merged ranges separately.
        /// </summary>
        private const XLCellsUsedOptions CopiedCellsUsedOptions = XLCellsUsedOptions.All
                                                                  & ~XLCellsUsedOptions.ConditionalFormats
                                                                  & ~XLCellsUsedOptions.DataValidation
                                                                  & ~XLCellsUsedOptions.MergedRanges;

        /// <summary>
        /// Throw when a used cell or a merged range of the source has its copy after the last row or the
        /// last column of the sheet. Excel does not paste a copy that does not fit. Call it before the first
        /// change to the target.
        /// </summary>
        /// <param name="source">The copied range. This cell is the first cell of the copy.</param>
        /// <exception cref="InvalidOperationException">
        /// A used cell or a merged range of the source has its copy after the last row or the last column.
        /// </exception>
        private void ThrowIfCopyIsOutsideSheet(XLRangeBase source)
        {
            var sourceArea = source.SheetRange;
            var rowsInSheet = XLHelper.MaxRowNumber - _rowNumber + 1;
            var columnsInSheet = XLHelper.MaxColumnNumber - _columnNumber + 1;
            if (sourceArea.Height > rowsInSheet)
                ThrowIfUsed(source.Worksheet, new Area(sourceArea.TopRow + rowsInSheet, sourceArea.LeftColumn, sourceArea.BottomRow, sourceArea.RightColumn));

            if (sourceArea.Width > columnsInSheet)
                ThrowIfUsed(source.Worksheet, new Area(sourceArea.TopRow, sourceArea.LeftColumn + columnsInSheet, sourceArea.BottomRow, sourceArea.RightColumn));

            // An empty merged cell is not a used cell, but the copy merges the cells of the target.
            foreach (var mergedRange in source.Worksheet.Internals.MergedRanges.Where<XLRange>(source.Contains))
            {
                var mergedArea = mergedRange.SheetRange;
                if (mergedArea.BottomRow - sourceArea.TopRow >= rowsInSheet || mergedArea.RightColumn - sourceArea.LeftColumn >= columnsInSheet)
                    throw CreateOutsideSheetException(source.Worksheet, mergedArea);
            }

            static void ThrowIfUsed(XLWorksheet sheet, Area outsideCopies)
            {
                var range = sheet.Range(outsideCopies.TopRow, outsideCopies.LeftColumn, outsideCopies.BottomRow, outsideCopies.RightColumn);
                if (range.CellsUsed(CopiedCellsUsedOptions).Any())
                    throw CreateOutsideSheetException(sheet, outsideCopies);
            }

            static InvalidOperationException CreateOutsideSheetException(XLWorksheet sheet, Area sourceCells)
            {
                return new InvalidOperationException($"The copy does not fit in the sheet. The source cells in {sheet.Name.EscapeSheetName()}!{sourceCells} have their copy after the last row or the last column.");
            }
        }

        private void ClearMerged()
        {
            List<IXLRange> mergeToDelete = Worksheet.Internals.MergedRanges.GetIntersectedRanges(Address).ToList();

            mergeToDelete.ForEach(m => Worksheet.Internals.MergedRanges.Remove(m));
        }

        internal string GetFormulaR1C1(string value)
        {
            return XLCellFormula.GetFormula(value, FormulaConversionType.A1ToR1C1, new Point(_rowNumber, _columnNumber));
        }

        internal string GetFormulaA1(string value)
        {
            return XLCellFormula.GetFormula(value, FormulaConversionType.R1C1ToA1, new Point(_rowNumber, _columnNumber));
        }

        internal void CopyValuesFrom(XLCell source)
        {
            // Rich text is basically a super set of a value. Setting a value would override rich text and vice versa.
            var sourceRichText = source.SliceRichText;
            if (sourceRichText is null)
                SliceCellValue = source.SliceCellValue;
            else
                SliceRichText = sourceRichText;

            // An array formula and a data table have one formula for many cells. The caller copies them separately.
            FormulaR1C1 = source.Formula is { Type: not FormulaType.Normal } ? null : source.FormulaR1C1;
            SliceComment = source.SliceComment == null ? null : XLComment.CreateAsCopy(this, source, source.SliceComment);

            if (source.Worksheet.Hyperlinks.TryGet(source.Point, out var sourceHyperlink))
            {
                SetCellHyperlink(new XLHyperlink(sourceHyperlink));
            }
        }

        private IXLCell GetTargetCell(String target, XLWorksheet defaultWorksheet)
        {
            var pair = target.Split('!');
            if (pair.Length == 1)
                return defaultWorksheet.Cell(target) ?? throw new ArgumentException($"'{target}' is not A1 address or named range.", nameof(target));

            var wsName = pair[0];
            if (wsName.StartsWith("'"))
                wsName = wsName.Substring(1, wsName.Length - 2);
            return defaultWorksheet.Workbook.Worksheet(wsName).Cell(pair[1]);
        }

        internal IXLCell CopyFromInternal(XLCell otherCell, XLCellCopyOptions options)
        {
            if (options.HasFlag(XLCellCopyOptions.Values))
                CopyValuesFrom(otherCell);

            // Other cell might be from a different workbook.
            if (options.HasFlag(XLCellCopyOptions.Styles))
                FormatValue = Styles.GetRegisteredCellFormat(otherCell.GetFormat());

            if (options.HasFlag(XLCellCopyOptions.Sparklines))
                CopySparklineFrom(otherCell);

            if (options.HasFlag(XLCellCopyOptions.ConditionalFormats))
                Worksheet.ConditionalFormats.CopyFrom(otherCell.Worksheet, otherCell.Point, Point, true);

            if (options.HasFlag(XLCellCopyOptions.DataValidations))
                CopyDataValidationFrom(otherCell);

            return this;
        }

        private void CopySparklineFrom(XLCell otherCell)
        {
            if (otherCell.Sparkline is not { } otherSparkline) return;

            // A sparkline without source data has nothing to show, thus it is not copied.
            if (otherSparkline.SourceData is not XLRangeBase otherSourceData) return;

            var sourceDataAddress = otherSourceData.RangeAddress.ToString();
            var shiftedRangeAddress = GetFormulaA1(otherCell.GetFormulaR1C1(sourceDataAddress));
            var sourceDataWorksheet = otherCell.Worksheet == otherSourceData.Worksheet
                ? Worksheet
                : otherSourceData.Worksheet;
            var sourceData = sourceDataWorksheet.Range(shiftedRangeAddress) ?? throw new InvalidOperationException($"The shifted source data address '{shiftedRangeAddress}' is not valid.");

            IXLSparklineGroup group;
            if (otherCell.Worksheet == Worksheet)
            {
                group = otherSparkline.SparklineGroup;
            }
            else
            {
                group = Worksheet.SparklineGroups.Add(new XLSparklineGroup(Worksheet, otherSparkline.SparklineGroup));
                if (otherSparkline.SparklineGroup.DateRange is XLRangeBase otherDateRange)
                {
                    var dateRangeWorksheet =
                        otherCell.Worksheet == otherDateRange.Worksheet
                            ? Worksheet
                            : otherDateRange.Worksheet;
                    var dateRangeAddress = otherDateRange.RangeAddress.ToString();
                    var shiftedDateRangeAddress = GetFormulaA1(otherCell.GetFormulaR1C1(dateRangeAddress));
                    group.SetDateRange(dateRangeWorksheet.Range(shiftedDateRangeAddress));
                }
            }

            group.Add(this, sourceData);
        }

        public IXLCell CopyFrom(IXLCell otherCell, XLCellCopyOptions options)
        {
            var source = (XLCell)otherCell; // To expose GetFormulaR1C1, etc
            if (!options.HasFlag(XLCellCopyOptions.Values))
            {
                CopyFromInternal(source, options);
                return this;
            }

            // Excel does not allow a copy of only part of an array formula.
            var arrayFormula = source.Formula is { Type: FormulaType.Array } sourceFormula ? sourceFormula : null;
            if (arrayFormula is not null && arrayFormula.Range != new Area(source.Point))
                throw ArrayFormulaCopy.CreatePartialCopyException(arrayFormula.Range);

            if (arrayFormula is not null)
                ArrayFormulaCopy.ThrowIfMerged(this);

            // The copy replaces the contents of this cell. Excel does not allow a change of only part of an array formula.
            ArrayFormulaShiftCheck.ThrowIfClearSplits(this, XLClearOptions.Contents);

            // Create the copy of the array formula before the first change, because the parse can throw.
            var arrayFormulaCopy = arrayFormula?.GetArrayCopy(new Area(Point));
            CopyFromInternal(source, options);
            if (arrayFormulaCopy is not null)
                Worksheet.SetArrayFormula(arrayFormulaCopy);

            return this;
        }

        private void CopyDataValidationFrom(XLCell otherCell)
        {
            if (otherCell.HasDataValidation)
                CopyDataValidation(otherCell, otherCell.GetDataValidation());
            else if (HasDataValidation)
            {
                Worksheet.DataValidations.Delete(new Area(Point));
            }
        }

        internal void CopyDataValidation(XLCell otherCell, IXLDataValidation otherDv)
        {
            var thisDv = (XLDataValidation)GetDataValidation();
            thisDv.CopyFrom(otherDv);
            thisDv.Value = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.Value));
            thisDv.MinValue = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MinValue));
            thisDv.MaxValue = GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MaxValue));
        }

        /// <summary>
        /// Shift the references in the formula of the cell, after an insert or a delete of cells.
        /// </summary>
        /// <param name="refMod">The visitor that shifts each reference.</param>
        /// <param name="shiftedSharedFormulas">The array formulas and data tables that the shift
        /// already changed. The method adds the formula of the cell to it.</param>
        internal void ShiftFormulaReferences(CopyVisitor refMod, HashSet<XLCellFormula> shiftedSharedFormulas)
        {
            if (Formula is not { } formula)
                return;

            if (formula.Type != FormulaType.Normal)
            {
                // The cells of an array formula or a data table share one formula. Shift it only
                // one time. The shift changes the range, so the range cannot identify the first cell.
                if (shiftedSharedFormulas.Add(formula))
                    formula.ShiftReferences(Worksheet.Name, refMod);

                return;
            }

            var formulaA1 = formula.A1;
            string shiftedFormulaA1;
            try
            {
                shiftedFormulaA1 = FormulaConverter.ModifyA1(formulaA1, Worksheet.Name, _rowNumber, _columnNumber, refMod);
            }
            catch (ParsingException)
            {
                // The parser cannot find the references of an unparsable formula. The formula stays the same.
                shiftedFormulaA1 = formulaA1;
            }

            // Do not create a new formula without a change. The setter also marks the formula as
            // dirty, so do that here too.
            if (shiftedFormulaA1 == formulaA1)
            {
                InvalidateFormula();
                return;
            }

            FormulaA1 = shiftedFormulaA1;
        }

        private XLCell CellShift(Int32 rowsToShift, Int32 columnsToShift)
        {
            return Worksheet.Cell(_rowNumber + rowsToShift, _columnNumber + columnsToShift);
        }

        #region XLCell Above

        IXLCell IXLCell.CellAbove()
        {
            return CellAbove();
        }

        IXLCell IXLCell.CellAbove(Int32 step)
        {
            return CellAbove(step);
        }

        public XLCell CellAbove()
        {
            return CellAbove(1);
        }

        public XLCell CellAbove(Int32 step)
        {
            return CellShift(step * -1, 0);
        }

        #endregion XLCell Above

        #region XLCell Below

        IXLCell IXLCell.CellBelow()
        {
            return CellBelow();
        }

        IXLCell IXLCell.CellBelow(Int32 step)
        {
            return CellBelow(step);
        }

        public XLCell CellBelow()
        {
            return CellBelow(1);
        }

        public XLCell CellBelow(Int32 step)
        {
            return CellShift(step, 0);
        }

        #endregion XLCell Below

        #region XLCell Left

        IXLCell IXLCell.CellLeft()
        {
            return CellLeft();
        }

        IXLCell IXLCell.CellLeft(Int32 step)
        {
            return CellLeft(step);
        }

        public XLCell CellLeft()
        {
            return CellLeft(1);
        }

        public XLCell CellLeft(Int32 step)
        {
            return CellShift(0, step * -1);
        }

        #endregion XLCell Left

        #region XLCell Right

        IXLCell IXLCell.CellRight()
        {
            return CellRight();
        }

        IXLCell IXLCell.CellRight(Int32 step)
        {
            return CellRight(step);
        }

        public XLCell CellRight()
        {
            return CellRight(1);
        }

        public XLCell CellRight(Int32 step)
        {
            return CellShift(0, step);
        }

        #endregion XLCell Right

        [MemberNotNullWhen(true, nameof(Formula))]
        public Boolean HasFormula => Formula is not null;

        public Boolean HasArrayFormula => Formula?.Type == FormulaType.Array;

        public IXLRangeAddress? FormulaReference
        {
            get
            {
                if (Formula is null)
                    return null;

                // An array formula and a data table always have a range. The default range is A1:A1,
                // so a comparison with the default range is correct only for a normal formula.
                var range = Formula.Range;
                if (Formula.Type == FormulaType.Normal && range == default)
                    return null;

                return XLRangeAddress.FromSheetRange(Worksheet, range);
            }
            set
            {
                if (Formula is null)
                    throw new ArgumentException("Cell doesn't contain a formula.");

                if (value is null)
                {
                    if (Formula.Type != FormulaType.Normal)
                        throw new ArgumentException("An array formula or a data table must have a range. To remove the formula, set the value of its range.");

                    Formula.Range = default;
                    return;
                }

                if (value.Worksheet is not null && Worksheet != value.Worksheet)
                    throw new ArgumentException("The reference worksheet must be same as worksheet of the cell or null.");

                Formula.Range = Area.FromRangeAddress(value);
            }
        }

        public IXLRange CurrentRegion => Worksheet.Range(FindCurrentRegion());

        private IXLRangeAddress FindCurrentRegion()
        {
            var sheet = Worksheet;

            var minRow = _rowNumber;
            var minCol = _columnNumber;
            var maxRow = _rowNumber;
            var maxCol = _columnNumber;

            bool hasRegionExpanded;

            do
            {
                hasRegionExpanded = false;

                var borderMinRow = Math.Max(minRow - 1, XLHelper.MinRowNumber);
                var borderMaxRow = Math.Min(maxRow + 1, XLHelper.MaxRowNumber);
                var borderMinColumn = Math.Max(minCol - 1, XLHelper.MinColumnNumber);
                var borderMaxColumn = Math.Min(maxCol + 1, XLHelper.MaxColumnNumber);

                if (minCol > XLHelper.MinColumnNumber &&
                    !IsVerticalBorderBlank(sheet, borderMinColumn, borderMinRow, borderMaxRow))
                {
                    hasRegionExpanded = true;
                    minCol = borderMinColumn;
                }

                if (maxCol < XLHelper.MaxColumnNumber &&
                    !IsVerticalBorderBlank(sheet, borderMaxColumn, borderMinRow, borderMaxRow))
                {
                    hasRegionExpanded = true;
                    maxCol = borderMaxColumn;
                }

                if (minRow > XLHelper.MinRowNumber &&
                    !IsHorizontalBorderBlank(sheet, borderMinRow, borderMinColumn, borderMaxColumn))
                {
                    hasRegionExpanded = true;
                    minRow = borderMinRow;
                }

                if (maxRow < XLHelper.MaxRowNumber &&
                    !IsHorizontalBorderBlank(sheet, borderMaxRow, borderMinColumn, borderMaxColumn))
                {
                    hasRegionExpanded = true;
                    maxRow = borderMaxRow;
                }
            } while (hasRegionExpanded);

            return new XLRangeAddress(
                new XLAddress(sheet, minRow, minCol, false, false),
                new XLAddress(sheet, maxRow, maxCol, false, false));

            static bool IsVerticalBorderBlank(XLWorksheet sheet, int borderColumn, int borderMinRow, int borderMaxRow)
            {
                for (var row = borderMinRow; row <= borderMaxRow; row++)
                {
                    var verticalBorderCell = sheet.Cell(row, borderColumn);
                    if (!verticalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
                    {
                        return false;
                    }
                }

                return true;
            }

            static bool IsHorizontalBorderBlank(XLWorksheet sheet, int borderRow, int borderMinColumn, int borderMaxColumn)
            {
                for (var col = borderMinColumn; col <= borderMaxColumn; col++)
                {
                    var horizontalBorderCell = sheet.Cell(borderRow, col);
                    if (!horizontalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        internal bool IsInferiorMergedCell()
        {
            return MergedRange() is { } mergedRange && !Address.Equals(mergedRange.RangeAddress.FirstAddress);
        }

        /// <summary>
        /// Get glyph bounding boxes for each grapheme in the text. Box size is determined according to
        /// the font of a grapheme. New lines are represented as default (all dimensions zero) box.
        /// A line without any text (i.e. contains only new line) should be represented by a box
        /// with zero advance width, but with a line height of corresponding font.
        /// </summary>
        /// <param name="engine">Engine used to determine box size.</param>
        /// <param name="dpi">DPI used to determine size of glyphs.</param>
        /// <param name="output">List where items are added.</param>
        internal void GetGlyphBoxes(IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
        {
            var richText = SliceRichText;
            if (richText is not null)
            {
                foreach (var richTextRun in richText.Runs)
                {
                    var text = richText.GetRunText(richTextRun);
                    var font = richTextRun.Font.ToFontBase();
                    AddGlyphs(text, font, engine, dpi, output);
                }
            }
            else
            {
                var text = GetFormattedString();
                AddGlyphs(text, Style.Font, engine, dpi, output);
            }

            static void AddGlyphs(string text, IXLFontBase font, IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
            {
                Span<int> zeroWidthJoiner = stackalloc int[1] { 0x200D };
                var prevWasNewLine = false;
                var graphemeStarts = StringInfo.ParseCombiningCharacters(text);
                var textSpan = text.AsSpan();

                // If we have more than 1 code unit per grapheme, the code units can
                // be distributed through multiple grapheme. In the worst case, all extra
                // code units are in exactly one grapheme -> allocate buffer of that size.
                Span<int> codePointsBuffer = stackalloc int[1 + text.Length - graphemeStarts.Length];
                for (var i = 0; i < graphemeStarts.Length; ++i)
                {
                    var startIdx = graphemeStarts[i];
                    var slice = textSpan.Slice(startIdx);
                    if (slice.TrySliceNewLine(out var eolLen))
                    {
                        i += eolLen - 1;
                        if (prevWasNewLine)
                        {
                            // If there are consecutive new lines, we need height of new the lines between them
                            var box = engine.GetGlyphBox(zeroWidthJoiner, font, dpi);
                            output.Add(box);
                        }

                        output.Add(GlyphBox.LineBreak);
                        prevWasNewLine = true;
                    }
                    else
                    {
                        var codeUnits = i + 1 < graphemeStarts.Length
                            ? textSpan.Slice(startIdx, graphemeStarts[i + 1] - startIdx)
                            : textSpan.Slice(startIdx);
                        var count = codeUnits.ToCodePoints(codePointsBuffer);
                        ReadOnlySpan<int> grapheme = codePointsBuffer.Slice(0, count);
                        var box = engine.GetGlyphBox(grapheme, font, dpi);
                        output.Add(box);
                        prevWasNewLine = false;
                    }
                }
            }
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Point.GetHashCode() * 397) ^ Worksheet.GetHashCode();
            }
        }

        public override bool Equals(object? obj)
        {
            return obj is XLCell cell && cell.Worksheet == Worksheet && cell.Point == Point;
        }
    }
}
