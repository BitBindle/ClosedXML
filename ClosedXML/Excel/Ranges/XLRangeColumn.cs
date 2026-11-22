using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeColumn : XLRangeBase, IXLRangeColumn
    {
        /// <summary>
        /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
        /// </summary>
        public XLRangeColumn(XLRangeAddress rangeAddress)
            : base(rangeAddress)
        {
        }

        internal override XLCellFormat Format => XLCellFormat.ForRange(Worksheet, RangeAddress);

        #region IXLRangeColumn Members

        IXLCell IXLRangeColumn.Cell(int rowNumber)
        {
            return Cell(rowNumber);
        }

        IXLCells IXLRangeColumn.Cells(string cellsInColumn) => Cells(cellsInColumn);

        public override XLCells Cells(string cellsInColumn)
        {
            var retVal = new XLCells(Worksheet, false, XLCellsUsedOptions.AllContents);
            var rangePairs = cellsInColumn.Split(',');
            foreach (string pair in rangePairs)
                retVal.Add(Range(pair.Trim()).RangeAddress);
            return retVal;
        }

        public IXLCells Cells(int firstRow, int lastRow)
        {
            return Cells(firstRow + ":" + lastRow);
        }

        public void Delete()
        {
            Delete(true);
        }

        internal void Delete(Boolean deleteTableField)
        {
            if (deleteTableField && IsTableColumn())
            {
                var table = Table;
                if (!Cell(1).Value.TryGetText(out var firstCellValue))
                    throw new InvalidOperationException("Top cell doesn't contain a text.");

                if (!table.FieldNames.ContainsKey(firstCellValue))
                    throw new InvalidOperationException($"Field {firstCellValue} not found.");

                var field = table.Fields.Cast<XLTableField>().Single(f => f.Name == firstCellValue);
                field.Delete(false);
            }

            Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }

        public IXLCells InsertCellsAbove(int numberOfRows)
        {
            return InsertCellsAbove(numberOfRows, false);
        }

        public IXLCells InsertCellsAbove(int numberOfRows, bool expandRange)
        {
            return InsertRowsAbove(numberOfRows, expandRange).Cells();
        }

        public IXLCells InsertCellsBelow(int numberOfRows)
        {
            return InsertCellsBelow(numberOfRows, true);
        }

        public IXLCells InsertCellsBelow(int numberOfRows, bool expandRange)
        {
            return InsertRowsBelow(numberOfRows, expandRange).Cells();
        }

        public int CellCount()
        {
            return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
        }

        public IXLRangeColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, Boolean matchCase = false, Boolean ignoreBlanks = true)
        {
            base.Sort(1, sortOrder, matchCase, ignoreBlanks);
            return this;
        }

        public IXLRangeColumn CopyTo(IXLCell target)
        {
            base.CopyTo((XLCell)target);

            int lastRowNumber = target.Address.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            int lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.Address.RowNumber,
                target.Address.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Column(1);
        }

        public new IXLRangeColumn CopyTo(IXLRangeBase target)
        {
            base.CopyTo(target);

            int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
            if (lastRowNumber > XLHelper.MaxRowNumber)
                lastRowNumber = XLHelper.MaxRowNumber;
            int lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
            if (lastColumnNumber > XLHelper.MaxColumnNumber)
                lastColumnNumber = XLHelper.MaxColumnNumber;

            return target.Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.FirstAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber)
                .Column(1);
        }

        public IXLRangeColumn Column(int start, int end)
        {
            return Range(start, end).FirstColumn();
        }

        public IXLRangeColumn Column(IXLCell start, IXLCell end)
        {
            return Column(start.Address.RowNumber, end.Address.RowNumber);
        }

        public IXLRangeColumns Columns(string columns)
        {
            var retVal = new XLRangeColumns(Worksheet);
            var rowPairs = columns.Split(',');
            foreach (string trimmedPair in rowPairs.Select(pair => pair.Trim()))
            {
                string firstRow;
                string lastRow;
                if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
                {
                    var rowRange = trimmedPair.Split(':', '-');

                    firstRow = rowRange[0];
                    lastRow = rowRange[1];
                }
                else
                {
                    firstRow = trimmedPair;
                    lastRow = trimmedPair;
                }

                retVal.Add(Range(firstRow + ":" + lastRow).FirstColumn());
            }

            return retVal;
        }

        public IXLColumn WorksheetColumn()
        {
            return Worksheet.Column(RangeAddress.FirstAddress.ColumnNumber);
        }

        public IXLRangeColumn AdjustToContents()
        {
            return AdjustToContents(RangeAddress.FirstAddress.RowNumber, RangeAddress.LastAddress.RowNumber);
        }

        public IXLRangeColumn AdjustToContents(Int32 startRow)
        {
            return AdjustToContents(startRow, RangeAddress.LastAddress.RowNumber);
        }

        public IXLRangeColumn AdjustToContents(Int32 startRow, Int32 endRow)
        {
            WorksheetColumn().AdjustToContents(startRow, endRow);
            return this;
        }

        public IXLRangeColumn AdjustToContents(Double minWidth, Double maxWidth)
        {
            return AdjustToContents(RangeAddress.FirstAddress.RowNumber, RangeAddress.LastAddress.RowNumber, minWidth, maxWidth);
        }

        public IXLRangeColumn AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            return AdjustToContents(startRow, RangeAddress.LastAddress.RowNumber, minWidth, maxWidth);
        }

        public IXLRangeColumn AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            WorksheetColumn().AdjustToContents(startRow, endRow, minWidth, maxWidth);
            return this;
        }

        #endregion IXLRangeColumn Members

        public override XLRangeType RangeType
        {
            get { return XLRangeType.RangeColumn; }
        }

        public XLCell Cell(int row)
        {
            return Cell(row, 1);
        }

        internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
        }

        internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
        {
            RangeAddress = (XLRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
        }

        public XLRange Range(int firstRow, int lastRow)
        {
            return Range(firstRow, 1, lastRow, 1);
        }

        public override XLRange Range(string rangeAddressStr)
        {
            string rangeAddressToUse;
            if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
            {
                if (rangeAddressStr.Contains('-'))
                    rangeAddressStr = rangeAddressStr.Replace('-', ':');

                var arrRange = rangeAddressStr.Split(':');
                string firstPart = arrRange[0];
                string secondPart = arrRange[1];
                rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
            }
            else
                rangeAddressToUse = FixColumnAddress(rangeAddressStr);

            var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
            return Range(rangeAddress);
        }

        private XLRangeColumn ColumnShift(Int32 columnsToShift)
        {
            Int32 columnNumber = ColumnNumber() + columnsToShift;
            return Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                columnNumber,
                RangeAddress.LastAddress.RowNumber,
                columnNumber).FirstColumn();
        }

        #region XLRangeColumn Left

        IXLRangeColumn IXLRangeColumn.ColumnLeft()
        {
            return ColumnLeft();
        }

        IXLRangeColumn IXLRangeColumn.ColumnLeft(Int32 step)
        {
            return ColumnLeft(step);
        }

        public XLRangeColumn ColumnLeft()
        {
            return ColumnLeft(1);
        }

        public XLRangeColumn ColumnLeft(Int32 step)
        {
            return ColumnShift(step * -1);
        }

        #endregion XLRangeColumn Left

        #region XLRangeColumn Right

        IXLRangeColumn IXLRangeColumn.ColumnRight()
        {
            return ColumnRight();
        }

        IXLRangeColumn IXLRangeColumn.ColumnRight(Int32 step)
        {
            return ColumnRight(step);
        }

        public XLRangeColumn ColumnRight()
        {
            return ColumnRight(1);
        }

        public XLRangeColumn ColumnRight(Int32 step)
        {
            return ColumnShift(step);
        }

        #endregion XLRangeColumn Right

        public IXLTable AsTable()
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().AsTable();
        }

        public IXLTable AsTable(string name)
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().AsTable(name);
        }

        public IXLTable CreateTable()
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().CreateTable();
        }

        public IXLTable CreateTable(string name)
        {
            if (IsTableColumn())
                throw new InvalidOperationException("This column is already part of a table.");

            return AsRange().CreateTable(name);
        }

        public new IXLRangeColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            base.Clear(clearOptions);
            return this;
        }


        public IXLRangeColumn? ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
        {
            var firstCell = FirstCellUsed(options);
            var lastCell = LastCellUsed(options);
            if (firstCell is null || lastCell is null)
                return null;

            return Column(firstCell, lastCell);
        }

        internal XLTable? Table { get; set; }

        [MemberNotNullWhen(true, nameof(Table))]
        public Boolean IsTableColumn()
        {
            return Table != null;
        }
    }
}
