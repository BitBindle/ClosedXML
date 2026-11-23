using System.Collections.Generic;
using ClosedXML.Excel.CalcEngine;

namespace ClosedXML.Excel
{
    internal class FormulaSlice : ISlice
    {
        private readonly XLWorksheet _sheet;
        private readonly XLCalcEngine _engine;
        private readonly Slice<XLCellFormula?> _formulas = new();

        public FormulaSlice(XLWorksheet sheet)
        {
            _sheet = sheet;
            _engine = sheet.Workbook.CalcEngine;
        }

        public bool IsEmpty => _formulas.IsEmpty;

        /// <summary>
        /// The data table formulas that were set in the sheet. The list can have formulas that the sheet
        /// does not have anymore. <see cref="GetDataTables"/> removes them.
        /// </summary>
        private readonly HashSet<XLCellFormula> _dataTables = new(ReferenceEqualityComparer.Instance);

        /// <summary>
        /// A cached delegate for <see cref="IsRemovedDataTable"/>, so that <see cref="GetDataTables"/> does not allocate.
        /// </summary>
        private System.Predicate<XLCellFormula>? _isRemovedDataTable;

        public int MaxColumn => _formulas.MaxColumn;

        public int MaxRow => _formulas.MaxRow;

        public Dictionary<int, int>.KeyCollection UsedColumns => _formulas.UsedColumns;

        public IEnumerable<int> UsedRows => _formulas.UsedRows;

        public void Clear(Area area)
        {
            _formulas.Clear(area);
        }

        public void DeleteAreaAndShiftLeft(Area areaToDelete)
        {
            _formulas.DeleteAreaAndShiftLeft(areaToDelete);
        }

        public void DeleteAreaAndShiftUp(Area areaToDelete)
        {
            _formulas.DeleteAreaAndShiftUp(areaToDelete);
        }

        public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false)
        {
            return _formulas.GetEnumerator(area, reverse);
        }

        public void InsertAreaAndShiftDown(Area areaToInsert)
        {
            _formulas.InsertAreaAndShiftDown(areaToInsert);
        }

        public void InsertAreaAndShiftRight(Area areaToInsert)
        {
            _formulas.InsertAreaAndShiftRight(areaToInsert);
        }

        public bool IsUsed(Point address)
        {
            return _formulas.IsUsed(address);
        }

        public void Swap(Point sp1, Point sp2)
        {
            var value1 = _formulas[sp1];
            var value2 = _formulas[sp2];

            value1 = value1?.GetMovedTo(sp1, sp2);
            value2 = value2?.GetMovedTo(sp2, sp1);

            Set(sp1, value2);
            Set(sp2, value1);
        }

        internal XLCellFormula? Get(Point point)
        {
            return _formulas[point];
        }

        /// <summary>
        /// Get the data table formulas of the sheet. A data table has its formula only in the first cell of
        /// its range.
        /// </summary>
        internal IReadOnlyCollection<XLCellFormula> GetDataTables()
        {
            if (_dataTables.Count == 0)
                return _dataTables;

            // An insert or a delete changes the range of a formula and moves the formula with its first cell.
            _dataTables.RemoveWhere(_isRemovedDataTable ??= IsRemovedDataTable);
            return _dataTables;
        }

        private bool IsRemovedDataTable(XLCellFormula dataTable)
        {
            return !ReferenceEquals(_formulas[dataTable.Range.FirstPoint], dataTable);
        }

        internal void Set(Point point, XLCellFormula? formula)
        {
            // Can't ref, because it is an alias for a memory and thus wouldn't hold old formula.
            var original = _formulas[point];
            if (ReferenceEquals(original, formula))
                return;

            _formulas.Set(point, formula);
            if (formula?.Type == FormulaType.DataTable)
                _dataTables.Add(formula);

            // Remove first, so calc chain doesn't choke on two formulas
            // in one cell when changing a formula of a cell.
            var bookPoint = new SheetPoint(_sheet.Name, point);
            if (original is not null)
                _engine.RemoveFormula(bookPoint, original);

            if (formula is not null)
                _engine.AddNormalFormula(bookPoint, _sheet.Name, formula, _sheet.Workbook);
        }

        /// <summary>
        /// Set all cells in a <paramref name="range"/> to the array formula.
        /// </summary>
        /// <remarks>
        /// This method doesn't check that formula doesn't damage other array formulas.
        /// </remarks>
        internal void SetArray(Area range, XLCellFormula? arrayFormula)
        {
            for (var row = range.TopRow; row <= range.BottomRow; ++row)
            {
                for (var col = range.LeftColumn; col <= range.RightColumn; ++col)
                {
                    var point = new Point(row, col);
                    var original = _formulas[point];

                    _formulas.Set(point, arrayFormula);

                    // The formula removal removes formula from dependency tree
                    // (number of cells formula affects doesn't matter) and also
                    // removes point from the calc chain. Therefore, it works for
                    // array and normal formulas.
                    var bookPoint = new SheetPoint(_sheet.Name, point);
                    if (original is not null)
                        _engine.RemoveFormula(bookPoint, original);
                }
            }

            if (arrayFormula is not null)
                _engine.AddArrayFormula(range, arrayFormula, _sheet);
        }

        /// <summary>
        /// Get an enumerator of the cells with a formula. A cell without a formula is not enumerated,
        /// but the type of the value is nullable, because the slice can store null.
        /// </summary>
        internal Slice<XLCellFormula?>.Enumerator GetForwardEnumerator(Area range)
        {
            return new Slice<XLCellFormula?>.Enumerator(_formulas, range);
        }

        /// <summary>
        /// Mark all formulas in a range as dirty.
        /// </summary>
        internal void MarkDirty(Area range)
        {
            using var enumerator = GetForwardEnumerator(range);
            while (enumerator.MoveNext())
            {
                if (enumerator.Current is { } formula)
                    formula.IsDirty = true;
            }
        }
    }
}
