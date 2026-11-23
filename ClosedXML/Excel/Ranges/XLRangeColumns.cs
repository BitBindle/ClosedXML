using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLRangeColumns : IXLRangeColumns
    {
        private readonly XLWorksheet _worksheet;
        private readonly List<XLRangeColumn> _ranges = new List<XLRangeColumn>();

        public XLRangeColumns(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        internal XLCellFormat Format
        {
            get
            {
                var columns = _ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
                return XLCellFormat.ForAreas(_worksheet.Workbook, columns, null);
            }
        }

        #region IXLRangeColumns Members

        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        public IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            using (ArrayFormulaShiftCheck.BeginClear(_ranges, clearOptions))
                _ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Delete()
        {
            using (ArrayFormulaShiftCheck.BeginDeleteOfParts(_ranges, rows: false, byFirstLineNumber: false))
                _ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
        }

        public IXLRangeColumns AdjustToContents()
        {
            _ranges.ForEach(c => c.AdjustToContents());
            return this;
        }

        public IXLRangeColumns AdjustToContents(Int32 startRow)
        {
            _ranges.ForEach(c => c.AdjustToContents(startRow));
            return this;
        }

        public IXLRangeColumns AdjustToContents(Int32 startRow, Int32 endRow)
        {
            _ranges.ForEach(c => c.AdjustToContents(startRow, endRow));
            return this;
        }

        public IXLRangeColumns AdjustToContents(Double minWidth, Double maxWidth)
        {
            _ranges.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
            return this;
        }

        public IXLRangeColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth)
        {
            _ranges.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
            return this;
        }

        public IXLRangeColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth)
        {
            _ranges.ForEach(c => c.AdjustToContents(startRow, endRow, minWidth, maxWidth));
            return this;
        }

        public IXLRangeColumns Where(Func<IXLRangeColumn, Boolean> predicate)
        {
            if (predicate is null)
                throw new ArgumentNullException(nameof(predicate));

            return FromOrdered(OrderedColumns.Where(predicate));
        }

        public IXLRangeColumns Skip(Int32 count)
        {
            return FromOrdered(OrderedColumns.Skip(count));
        }

        public IXLRangeColumns Take(Int32 count)
        {
            return FromOrdered(OrderedColumns.Take(count));
        }

        public void Add(IXLRangeColumn range)
        {
            _ranges.Add((XLRangeColumn)range);
        }

        public IEnumerator<IXLRangeColumn> GetEnumerator()
        {
            return _ranges.Cast<IXLRangeColumn>()
              .OrderBy(r => r.Worksheet.Position)
              .ThenBy(r => r.ColumnNumber())
              .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(_worksheet, usedCellsOnly: false, options: XLCellsUsedOptions.AllContents);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(_worksheet, usedCellsOnly: true, options: XLCellsUsedOptions.AllContents);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }


        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(_worksheet, usedCellsOnly: true, options: options);
            foreach (XLRangeColumn container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        #endregion IXLRangeColumns Members

        /// <summary>
        /// The range columns in the order of the enumeration. The type makes a call to Where, Skip or Take use LINQ, not the method of this class.
        /// </summary>
        private IEnumerable<IXLRangeColumn> OrderedColumns => this;

        private IXLRangeColumns FromOrdered(IEnumerable<IXLRangeColumn> ranges)
        {
            var result = new XLRangeColumns(_worksheet);
            foreach (var range in ranges)
                result.Add(range);
            return result;
        }
    }
}
