using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal class XLTableRows : IXLTableRows
    {
        private readonly XLWorksheet _worksheet;
        private readonly List<XLTableRow> _ranges = new List<XLTableRow>();

        public XLTableRows(XLWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        #region IXLTableRows Members

        public IXLStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        public IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
        {
            using (ArrayFormulaShiftCheck.BeginClear(_ranges, clearOptions))
                _ranges.ForEach(r => r.Clear(clearOptions));
            return this;
        }

        public void Delete()
        {
            using (ArrayFormulaShiftCheck.BeginDeleteOfParts(_ranges, rows: true, byFirstLineNumber: false))
                _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
        }

        public void Add(IXLTableRow tableRow)
        {
            _ranges.Add((XLTableRow)tableRow);
        }

        public IEnumerator<IXLTableRow> GetEnumerator()
        {
            var retList = new List<IXLTableRow>();
            _ranges.ForEach(retList.Add);
            return retList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXLCells Cells()
        {
            var cells = new XLCells(_worksheet, false, XLCellsUsedOptions.AllContents);
            foreach (XLTableRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed()
        {
            var cells = new XLCells(_worksheet, true, XLCellsUsedOptions.AllContents);
            foreach (XLTableRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXLCells CellsUsed(Boolean includeFormats)
        {
            return CellsUsed(includeFormats
                ? XLCellsUsedOptions.All
                : XLCellsUsedOptions.AllContents);
        }

        public IXLCells CellsUsed(XLCellsUsedOptions options)
        {
            var cells = new XLCells(_worksheet, false, options);
            foreach (XLTableRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        #endregion IXLTableRows Members

        internal XLCellFormat Format
        {
            get
            {
                var rowAreas = _ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
                return XLCellFormat.ForTableRows(_worksheet, rowAreas);
            }
        }
    }
}
