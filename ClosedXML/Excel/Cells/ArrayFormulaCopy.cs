using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Copies array formulas and data tables. The copy of a cell copies the value of the cell. An
    /// array formula has one formula for all cells of its range, so the copy sets it after the copy of
    /// the cells.
    /// </summary>
    internal static class ArrayFormulaCopy
    {
        /// <summary>
        /// Get the array formulas that have cells in the copied area.
        /// </summary>
        /// <param name="sheet">The sheet of the copied cells.</param>
        /// <param name="area">The copied cells.</param>
        /// <returns>The array formulas. The list has each array formula one time.</returns>
        /// <exception cref="InvalidOperationException">The area has only part of an array formula.</exception>
        internal static List<XLCellFormula> GetArrayFormulas(XLWorksheet sheet, Area area)
        {
            var arrayFormulas = new List<XLCellFormula>();
            var foundFormulas = new HashSet<XLCellFormula>(ReferenceEqualityComparer.Instance);
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(area);
            while (enumerator.MoveNext())
            {
                if (enumerator.Current is not { Type: FormulaType.Array } formula || !foundFormulas.Add(formula))
                    continue;

                // Excel does not allow a copy of only part of an array formula.
                if (!area.Covers(formula.Range))
                    throw CreatePartialCopyException(formula.Range);

                arrayFormulas.Add(formula);
            }

            return arrayFormulas;
        }

        /// <summary>
        /// Create the copies of the array formulas in a copied area. Call it before the first change to
        /// the target, because it throws when the copy is not possible.
        /// </summary>
        /// <param name="source">The sheet of the copied cells.</param>
        /// <param name="sourceArea">The copied cells.</param>
        /// <param name="target">The sheet of the copy.</param>
        /// <param name="targetOrigin">The first cell of the copy.</param>
        /// <returns>The array formulas for the target. Give them to <see cref="SetArrayFormulas"/> after the copy of the cells.</returns>
        /// <exception cref="InvalidOperationException">
        /// The area has only part of an array formula, or a copy with more than one cell is over a table of the target.
        /// </exception>
        internal static List<XLCellFormula> CreateCopies(XLWorksheet source, Area sourceArea, XLWorksheet target, Point targetOrigin)
        {
            var arrayFormulas = GetArrayFormulas(source, sourceArea);
            var rowShift = targetOrigin.Row - sourceArea.TopRow;
            var columnShift = targetOrigin.Column - sourceArea.LeftColumn;
            var copies = new List<XLCellFormula>(arrayFormulas.Count);
            foreach (var arrayFormula in arrayFormulas)
            {
                // The copy of the cells throws for a cell outside the sheet, so this copy is not necessary.
                var range = arrayFormula.Range;
                var lastRow = range.BottomRow + rowShift;
                var lastColumn = range.RightColumn + columnShift;
                if (lastRow > XLHelper.MaxRowNumber || lastColumn > XLHelper.MaxColumnNumber)
                    continue;

                var targetRange = new Area(range.TopRow + rowShift, range.LeftColumn + columnShift, lastRow, lastColumn);

                // Excel allows an array formula in one cell of a table.
                if ((targetRange.Width > 1 || targetRange.Height > 1) && target.Tables.Any<XLTable>(table => Area.FromRangeAddress(table.RangeAddress).Intersects(targetRange)))
                    throw new InvalidOperationException($"The copy puts the array formula in {targetRange} over a table. A table cannot contain an array formula with more than one cell.");

                copies.Add(arrayFormula.GetArrayCopy(targetRange));
            }

            return copies;
        }

        /// <summary>
        /// Set the copies of the array formulas in the target sheet.
        /// </summary>
        /// <param name="target">The sheet of the copy.</param>
        /// <param name="copies">The array formulas from <see cref="CreateCopies"/>.</param>
        internal static void SetArrayFormulas(XLWorksheet target, List<XLCellFormula> copies)
        {
            foreach (var copy in copies)
                target.SetArrayFormula(copy);
        }

        /// <summary>
        /// Throw when a cell copy puts an array formula in a merged cell. Excel does not allow an array
        /// formula in merged cells, also in one cell.
        /// </summary>
        /// <param name="target">The cell of the copy.</param>
        /// <exception cref="InvalidOperationException">The cell is merged.</exception>
        internal static void ThrowIfMerged(XLCell target)
        {
            if (target.IsMerged())
                throw new InvalidOperationException($"The copy puts an array formula in the merged cell {target.Address.ToStringRelative()}. Merged cells cannot contain an array formula.");
        }

        /// <summary>
        /// Set a copy of each array formula and each data table of a sheet in a copy of the sheet. The
        /// cells of the copy have the same addresses as the cells of the source.
        /// </summary>
        /// <param name="source">The copied sheet.</param>
        /// <param name="target">The copy of the sheet.</param>
        internal static void CopyAll(XLWorksheet source, XLWorksheet target)
        {
            foreach (var arrayFormula in GetArrayFormulas(source, Area.Full))
                target.SetArrayFormula(arrayFormula.GetCopy());

            foreach (var dataTable in source.Internals.CellsCollection.FormulaSlice.GetDataTables())
                target.SetDataTable(dataTable.GetCopy());
        }

        /// <summary>
        /// Create the exception for a copy of only part of an array formula.
        /// </summary>
        /// <param name="range">The range of the array formula.</param>
        internal static InvalidOperationException CreatePartialCopyException(Area range)
        {
            return new InvalidOperationException($"The copy has only part of the array formula in {range}. Copy all cells of the range.");
        }
    }
}
