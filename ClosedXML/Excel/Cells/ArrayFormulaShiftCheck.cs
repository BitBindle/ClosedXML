using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Stops an insert, a delete or a clear that changes only part of an array formula or a data table.
    /// Excel does not allow these operations, because the cells of an array formula must stay
    /// together as one rectangle.
    /// </summary>
    /// <remarks>
    /// The check uses lines and a cross span. For rows, a line is a row and the cross span is a
    /// span of columns. For columns, a line is a column and the cross span is a span of rows.
    /// </remarks>
    internal static class ArrayFormulaShiftCheck
    {
        /// <summary>
        /// Throw when an insert moves only part of an array formula.
        /// </summary>
        /// <param name="sheet">The sheet of the insert.</param>
        /// <param name="rows"><c>true</c> when the insert moves cells down, <c>false</c> when it moves cells right.</param>
        /// <param name="firstLine">The first line that moves.</param>
        /// <param name="numberOfLines">The number of inserted lines.</param>
        /// <param name="crossFirst">The first line of the cross span that moves.</param>
        /// <param name="crossLast">The last line of the cross span that moves.</param>
        /// <exception cref="InvalidOperationException">The insert moves only part of an array formula.</exception>
        internal static void ThrowIfInsertSplits(XLWorksheet sheet, bool rows, int firstLine, int numberOfLines, int crossFirst, int crossLast)
        {
            if (sheet.ArrayFormulaChecksSuspended > 0)
                return;

            var ranges = GetMovableRanges(sheet, rows, firstLine, crossFirst, crossLast);
            ApplyInsert(ranges, rows, firstLine, numberOfLines, crossFirst, crossLast);
        }

        /// <summary>
        /// Check an insert of full rows or full columns. Then suspend the check until you dispose the
        /// returned object, because the insert of the range inside checks the same lines again.
        /// </summary>
        /// <param name="sheet">The sheet of the insert.</param>
        /// <param name="rows"><c>true</c> for an insert of rows, <c>false</c> for an insert of columns.</param>
        /// <param name="firstLine">The first line that moves.</param>
        /// <param name="numberOfLines">The number of inserted lines.</param>
        /// <returns>An object that resumes the check when you dispose it.</returns>
        /// <exception cref="InvalidOperationException">The insert moves only part of an array formula.</exception>
        internal static IDisposable BeginInsertOfLines(XLWorksheet sheet, bool rows, int firstLine, int numberOfLines)
        {
            // The insert inside throws for an incorrect number of lines. Do not throw a different exception first.
            var lastLine = rows ? XLHelper.MaxRowNumber : XLHelper.MaxColumnNumber;
            if (numberOfLines <= 0 || numberOfLines > lastLine)
                return SuspendChecks(sheet);

            // The range insert moves the last line when the first line is after the sheet. Check the same line.
            var crossLast = rows ? XLHelper.MaxColumnNumber : XLHelper.MaxRowNumber;
            ThrowIfInsertSplits(sheet, rows, Math.Min(firstLine, lastLine), numberOfLines, 1, crossLast);
            return SuspendChecks(sheet);
        }

        /// <summary>
        /// Throw when a delete moves or deletes only part of an array formula.
        /// </summary>
        /// <param name="sheet">The sheet of the delete.</param>
        /// <param name="rows"><c>true</c> when the delete moves cells up, <c>false</c> when it moves cells left.</param>
        /// <param name="deleted">The deleted cells.</param>
        /// <exception cref="InvalidOperationException">The delete changes only part of an array formula.</exception>
        internal static void ThrowIfDeleteSplits(XLWorksheet sheet, bool rows, Area deleted)
        {
            if (sheet.ArrayFormulaChecksSuspended > 0)
                return;

            var span = Span.Of(deleted, rows);
            var ranges = GetMovableRanges(sheet, rows, span.FirstLine, span.CrossFirst, span.CrossLast);
            ApplyDelete(ranges, rows, span);
            ThrowIfPartlyDeleted(ranges);
        }

        /// <summary>
        /// Start a check of an operation with many steps. Add all steps in the order of the operation.
        /// Then call <see cref="OperationCheck.Begin"/> before the first change to the sheet.
        /// </summary>
        /// <param name="sheet">The sheet of the operation.</param>
        /// <returns>The check of the operation.</returns>
        internal static OperationCheck CheckOperation(XLWorksheet sheet)
        {
            return new OperationCheck(sheet);
        }

        /// <summary>
        /// Check a delete of many parts before the first part changes the sheet. Then suspend the check of
        /// each part until you dispose the returned object. A part alone can split an array formula that
        /// the full delete removes.
        /// </summary>
        /// <param name="parts">The parts to delete.</param>
        /// <param name="rows"><c>true</c> when the parts are rows, <c>false</c> when the parts are columns.</param>
        /// <param name="byFirstLineNumber">
        /// <c>true</c> when the caller deletes each part at the line numbers that the part has before the delete.
        /// <c>false</c> when the caller deletes the range objects, because the sheet moves a range object after each delete.
        /// </param>
        /// <returns>An object that resumes the check when you dispose it, or <c>null</c> for one part.</returns>
        /// <exception cref="InvalidOperationException">The delete changes only part of an array formula.</exception>
        internal static IDisposable? BeginDeleteOfParts(IReadOnlyList<XLRangeBase> parts, bool rows, bool byFirstLineNumber)
        {
            // An earlier delete can make a range object not valid. The delete of that part does not change the cells.
            var validParts = parts.Where(static part => part.RangeAddress.IsValid).ToList();
            if (validParts.Count < 2)
                return null;

            // Check all sheets before the first suspend.
            var sheets = new List<XLWorksheet>();
            foreach (var sheetParts in validParts.GroupBy(static part => part.Worksheet))
            {
                if (sheetParts.Key.ArrayFormulaChecksSuspended == 0)
                    CheckDeleteOfSheetParts(sheetParts.Key, sheetParts.Select(part => Span.Of(Area.FromRangeAddress(part.RangeAddress), rows)).ToList(), rows, byFirstLineNumber);

                sheets.Add(sheetParts.Key);
            }

            return new CheckResumers(sheets.Select(SuspendChecks).ToList());
        }

        /// <summary>
        /// Throw when a clear of one cell removes the contents of only part of an array formula or a data
        /// table.
        /// </summary>
        /// <remarks>
        /// A clear of one cell does not clear other cells, so the check does not suspend the checks of the
        /// sheet. The check reads only the formula of the cell and the list of data tables.
        /// </remarks>
        /// <param name="cell">The cleared cell.</param>
        /// <param name="clearOptions">The options of the clear.</param>
        /// <exception cref="InvalidOperationException">The clear removes the contents of only part of an array formula or a data table.</exception>
        internal static void ThrowIfClearSplits(XLCell cell, XLClearOptions clearOptions)
        {
            var sheet = cell.Worksheet;
            if (!clearOptions.HasFlag(XLClearOptions.Contents) || sheet.ArrayFormulaChecksSuspended > 0)
                return;

            var cellArea = new Area(cell.Point);
            var formulas = sheet.Internals.CellsCollection.FormulaSlice;
            if (formulas.Get(cell.Point) is { Type: FormulaType.Array } formula && formula.Range != cellArea)
                throw CreateException(formula.Range);

            foreach (var dataTable in formulas.GetDataTables())
            {
                if (dataTable.Range.Contains(cell.Point) && dataTable.Range != cellArea)
                    throw CreateException(dataTable.Range);
            }
        }

        /// <summary>
        /// Check that a clear of cells does not remove the contents of only part of an array formula or a
        /// data table. Then suspend the check until you dispose the returned object.
        /// </summary>
        /// <param name="cells">The cleared cells.</param>
        /// <param name="clearOptions">The options of the clear.</param>
        /// <returns>An object that resumes the check when you dispose it. <c>null</c> when the clear keeps the contents.</returns>
        /// <exception cref="InvalidOperationException">The clear removes the contents of only part of an array formula or a data table.</exception>
        internal static IDisposable? BeginClear(IEnumerable<XLCell> cells, XLClearOptions clearOptions)
        {
            return BeginClear(cells.Select(static cell => (cell.Worksheet, GetClearedArea(cell))), clearOptions);
        }

        /// <summary>
        /// Get the area that a clear of the cell changes. The clear of a merged cell clears all cells of the merged range.
        /// </summary>
        private static Area GetClearedArea(XLCell cell)
        {
            if (cell.IsMerged() && cell.Worksheet.Internals.MergedRanges.GetIntersectedRanges(cell.Address).FirstOrDefault() is { } mergedRange)
                return Area.FromRangeAddress(mergedRange.RangeAddress);

            return new Area(cell.Point);
        }

        /// <summary>
        /// Check that a clear of ranges does not remove the contents of only part of an array formula or a
        /// data table. Then suspend the check until you dispose the returned object.
        /// </summary>
        /// <param name="parts">The cleared ranges.</param>
        /// <param name="clearOptions">The options of the clear.</param>
        /// <returns>An object that resumes the check when you dispose it. <c>null</c> when the clear keeps the contents.</returns>
        /// <exception cref="InvalidOperationException">The clear removes the contents of only part of an array formula or a data table.</exception>
        internal static IDisposable? BeginClear(IEnumerable<XLRangeBase> parts, XLClearOptions clearOptions)
        {
            // An earlier delete can make a range object not valid. The clear of that part does not change the cells.
            var validParts = parts
                .Where(static part => part.RangeAddress.IsValid)
                .Select(static part => (part.Worksheet, Area.FromRangeAddress(part.RangeAddress)));
            return BeginClear(validParts, clearOptions);
        }

        /// <summary>
        /// Check all parts of a clear before the first change. The parts together must contain all cells
        /// of each array formula and data table that they change.
        /// </summary>
        private static IDisposable? BeginClear(IEnumerable<(XLWorksheet Sheet, Area Area)> parts, XLClearOptions clearOptions)
        {
            if (!clearOptions.HasFlag(XLClearOptions.Contents))
                return null;

            // Check all sheets before the first suspend.
            var sheets = new List<XLWorksheet>();
            foreach (var sheetParts in parts.GroupBy(static part => part.Sheet))
            {
                if (sheetParts.Key.ArrayFormulaChecksSuspended == 0)
                    ThrowIfClearSplits(sheetParts.Key, sheetParts.Select(static part => part.Area).ToList());

                sheets.Add(sheetParts.Key);
            }

            return new CheckResumers(sheets.Select(SuspendChecks).ToList());
        }

        private static void ThrowIfClearSplits(XLWorksheet sheet, List<Area> cleared)
        {
            var ranges = GetArrayRanges(sheet, cleared);
            if (ranges.Count == 0)
                return;

            // A clear of a cell collection gives one area for each cell. A set of these cells makes the
            // check of each range fast.
            var cells = new HashSet<Point>();
            var areas = new List<Area>();
            foreach (var area in cleared)
            {
                if (area.Width == 1 && area.Height == 1)
                    cells.Add(area.FirstPoint);
                else
                    areas.Add(area);
            }

            foreach (var range in ranges)
            {
                if (!ContainAllCells(cells, areas, range))
                    throw CreateException(range);
            }
        }

        /// <summary>
        /// Get <c>true</c> when the <paramref name="cells"/> and the <paramref name="areas"/> together contain
        /// all cells of the <paramref name="range"/>.
        /// </summary>
        private static bool ContainAllCells(HashSet<Point> cells, List<Area> areas, Area range)
        {
            if (areas.Exists(area => area.Intersect(range) == range))
                return true;

            foreach (var point in range)
            {
                if (!cells.Contains(point) && !areas.Exists(area => area.Contains(point)))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Suspend the check on the sheet until you dispose the returned object. Use it only after a check
        /// of all lines of the operation.
        /// </summary>
        /// <param name="sheet">The sheet of the operation.</param>
        /// <returns>An object that resumes the check when you dispose it.</returns>
        internal static IDisposable SuspendChecks(XLWorksheet sheet)
        {
            sheet.ArrayFormulaChecksSuspended++;
            return new CheckResumer(sheet);
        }

        private static void CheckDeleteOfSheetParts(XLWorksheet sheet, List<Span> parts, bool rows, bool byFirstLineNumber)
        {
            var firstLine = parts.Min(static part => part.FirstLine);
            var crossFirst = parts.Min(static part => part.CrossFirst);
            var crossLast = parts.Max(static part => part.CrossLast);
            var ranges = GetMovableRanges(sheet, rows, firstLine, crossFirst, crossLast);
            if (ranges.Count == 0)
                return;

            // The callers delete the parts from the last line to the first line. A stable sort keeps the order of equal parts.
            var orderedParts = parts.OrderByDescending(static part => part.FirstLine).ToList();

            // A delete at the line numbers from before the delete is easy to repeat on the ranges. A delete of range
            // objects is also easy to repeat when the parts have the same cross span and do not overlap. Then a delete
            // does not move the parts before it.
            if (byFirstLineNumber || HaveSameCrossAndDoNotOverlap(orderedParts))
            {
                foreach (var part in orderedParts)
                    ApplyDelete(ranges, rows, part);

                ThrowIfPartlyDeleted(ranges);
                return;
            }

            // The sheet moves and changes the other range objects after each delete. This check does not repeat that,
            // so it does not allow an array formula in the cells that can change. A part can change only the cells in
            // its cross span, from its first line to the end of the sheet.
            var changedRange = ranges.Find(range => parts.Any(part => CanChange(part, Span.Of(range.Current, rows))));
            if (changedRange is not null)
                throw CreateException(changedRange.Range);
        }

        private static bool CanChange(Span part, Span range)
        {
            return range.CrossFirst <= part.CrossLast && range.CrossLast >= part.CrossFirst && range.LastLine >= part.FirstLine;
        }

        private static bool HaveSameCrossAndDoNotOverlap(List<Span> orderedParts)
        {
            for (var i = 1; i < orderedParts.Count; ++i)
            {
                var previous = orderedParts[i - 1];
                var part = orderedParts[i];
                if (part.CrossFirst != previous.CrossFirst || part.CrossLast != previous.CrossLast || part.LastLine >= previous.FirstLine)
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Change the ranges in the same way as one delete changes the sheet. A range can lose some lines
        /// in one delete and the remaining lines in a later delete. For this reason, the check of lost lines is at the end.
        /// </summary>
        /// <exception cref="InvalidOperationException">The delete moves only some cells of a line of a range.</exception>
        private static void ApplyDelete(List<MovableRange> ranges, bool rows, Span deleted)
        {
            var deletedCount = deleted.LastLine - deleted.FirstLine + 1;
            for (var i = ranges.Count - 1; i >= 0; --i)
            {
                var range = ranges[i];
                var span = Span.Of(range.Current, rows);
                var isBefore = span.LastLine < deleted.FirstLine;
                var isOutsideCross = span.CrossLast < deleted.CrossFirst || span.CrossFirst > deleted.CrossLast;
                if (isBefore || isOutsideCross)
                    continue;

                if (!span.CrossIsInside(deleted.CrossFirst, deleted.CrossLast))
                    throw CreateException(range.Range);

                var overlapFirst = Math.Max(span.FirstLine, deleted.FirstLine);
                var overlapLast = Math.Min(span.LastLine, deleted.LastLine);
                var lostCount = Math.Max(0, overlapLast - overlapFirst + 1);
                if (lostCount == span.LastLine - span.FirstLine + 1)
                {
                    ranges.RemoveAt(i);
                    continue;
                }

                // The lines after the deleted lines move up. The lines before the deleted lines stay.
                var newFirst = span.FirstLine < deleted.FirstLine ? span.FirstLine : Math.Max(span.FirstLine - deletedCount, deleted.FirstLine);
                var newLast = newFirst + (span.LastLine - span.FirstLine - lostCount);
                range.Current = (span with { FirstLine = newFirst, LastLine = newLast }).ToArea(rows);
                range.LostLines |= lostCount > 0;
            }
        }

        /// <summary>
        /// Change the ranges in the same way as one insert changes the sheet.
        /// </summary>
        /// <exception cref="InvalidOperationException">The insert moves only part of a range.</exception>
        private static void ApplyInsert(List<MovableRange> ranges, bool rows, int firstLine, int numberOfLines, int crossFirst, int crossLast)
        {
            var maxLine = rows ? XLHelper.MaxRowNumber : XLHelper.MaxColumnNumber;
            foreach (var range in ranges)
            {
                var span = Span.Of(range.Current, rows);
                var isBefore = span.LastLine < firstLine;
                var isOutsideCross = span.CrossLast < crossFirst || span.CrossFirst > crossLast;
                if (isBefore || isOutsideCross)
                    continue;

                // The range moves as one rectangle only when it is fully in the cells that move. The insert
                // removes the cells that move out of the sheet, so the range must also stay in the sheet.
                if (span.FirstLine < firstLine || !span.CrossIsInside(crossFirst, crossLast) || (long)span.LastLine + numberOfLines > maxLine)
                    throw CreateException(range.Range);

                range.Current = (span with { FirstLine = span.FirstLine + numberOfLines, LastLine = span.LastLine + numberOfLines }).ToArea(rows);
            }
        }

        private static void ThrowIfPartlyDeleted(List<MovableRange> ranges)
        {
            foreach (var range in ranges)
            {
                if (range.LostLines)
                    throw CreateException(range.Range);
            }
        }

        private static List<MovableRange> GetMovableRanges(XLWorksheet sheet, bool rows, int firstLine, int crossFirst, int crossLast)
        {
            return GetMovableRanges(sheet, LinesToEnd(rows, firstLine, crossFirst, crossLast));
        }

        private static List<MovableRange> GetMovableRanges(XLWorksheet sheet, params IReadOnlyList<Area> areas)
        {
            return GetArrayRanges(sheet, areas)
                .Select(static range => new MovableRange(range))
                .ToList();
        }

        /// <summary>
        /// Get the cells in the cross span, from <paramref name="firstLine"/> to the end of the sheet.
        /// </summary>
        private static Area LinesToEnd(bool rows, int firstLine, int crossFirst, int crossLast)
        {
            return rows
                ? new Area(firstLine, crossFirst, XLHelper.MaxRowNumber, crossLast)
                : new Area(crossFirst, firstLine, crossLast, XLHelper.MaxColumnNumber);
        }

        /// <summary>
        /// Get the ranges of the array formulas and data tables that have cells in one of the
        /// <paramref name="areas"/>.
        /// </summary>
        /// <remarks>
        /// All cells of an array formula share one formula, so a search of the areas finds each array
        /// formula. A data table has the formula only in its first cell, and this cell can be outside
        /// the areas. The formula slice keeps a list of the data tables for this reason.
        /// </remarks>
        private static List<Area> GetArrayRanges(XLWorksheet sheet, IReadOnlyList<Area> areas)
        {
            var ranges = new List<Area>();
            var foundFormulas = new HashSet<XLCellFormula>(ReferenceEqualityComparer.Instance);
            var formulas = sheet.Internals.CellsCollection.FormulaSlice;
            foreach (var area in areas)
            {
                // A clear of a cell collection gives many areas with one cell. A direct read is faster than an enumerator.
                if (area.Width == 1 && area.Height == 1)
                {
                    if (formulas.Get(area.FirstPoint) is { Type: FormulaType.Array } cellFormula && foundFormulas.Add(cellFormula))
                        ranges.Add(cellFormula.Range);

                    continue;
                }

                using var enumerator = formulas.GetForwardEnumerator(area);
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is { Type: FormulaType.Array } formula && foundFormulas.Add(formula))
                        ranges.Add(formula.Range);
                }
            }

            foreach (var dataTable in formulas.GetDataTables())
            {
                if (IntersectsAny(dataTable.Range, areas))
                    ranges.Add(dataTable.Range);
            }

            return ranges;
        }

        private static bool IntersectsAny(Area range, IReadOnlyList<Area> areas)
        {
            foreach (var area in areas)
            {
                if (range.Intersects(area))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Throw when a table range has cells of an array formula or a data table with more than one cell.
        /// Excel does not allow these formulas in a table.
        /// </summary>
        /// <param name="sheet">The sheet of the table.</param>
        /// <param name="tableRange">The range of the table.</param>
        /// <exception cref="InvalidOperationException">The table range has cells of an array formula or a data table with more than one cell.</exception>
        internal static void ThrowIfTableHasArrayFormula(XLWorksheet sheet, Area tableRange)
        {
            if (FindArrayRangeWithMoreThanOneCell(sheet, tableRange) is { } range)
                throw new InvalidOperationException($"The table range has cells of the array formula or data table in {range}. A table cannot contain an array formula or a data table with more than one cell.");
        }

        /// <summary>
        /// Get <c>true</c> when the area has cells of an array formula or a data table with more than one cell.
        /// </summary>
        /// <param name="sheet">The sheet of the area.</param>
        /// <param name="area">The area to examine.</param>
        internal static bool HasArrayFormulaWithMoreThanOneCell(XLWorksheet sheet, Area area)
        {
            return FindArrayRangeWithMoreThanOneCell(sheet, area) is not null;
        }

        private static Area? FindArrayRangeWithMoreThanOneCell(XLWorksheet sheet, Area area)
        {
            foreach (var range in GetArrayRanges(sheet, [area]))
            {
                if (range.Width > 1 || range.Height > 1)
                    return range;
            }

            return null;
        }

        /// <summary>
        /// Throw when a merge range has cells of an array formula or a data table. Excel does not allow
        /// these formulas in merged cells, also in one cell.
        /// </summary>
        /// <param name="sheet">The sheet of the merge.</param>
        /// <param name="mergeRange">The range to merge.</param>
        /// <exception cref="InvalidOperationException">The merge range has cells of an array formula or a data table.</exception>
        internal static void ThrowIfMergeHasArrayFormula(XLWorksheet sheet, Area mergeRange)
        {
            var ranges = GetArrayRanges(sheet, [mergeRange]);
            if (ranges.Count > 0)
                throw new InvalidOperationException($"The merge range has cells of the array formula or data table in {ranges[0]}. Merged cells cannot contain an array formula or a data table.");
        }

        private static InvalidOperationException CreateException(Area range)
        {
            return new InvalidOperationException($"The operation changes only part of the array formula or data table in {range}. Change all cells of the range, or remove the formula first.");
        }

        private static InvalidOperationException CreateTransposeException(Area range)
        {
            return new InvalidOperationException($"The transpose moves cells of the array formula or data table in {range} to new positions. Remove the formula first.");
        }

        private readonly record struct Span(int FirstLine, int LastLine, int CrossFirst, int CrossLast)
        {
            internal static Span Of(Area area, bool rows)
            {
                return rows
                    ? new Span(area.TopRow, area.BottomRow, area.LeftColumn, area.RightColumn)
                    : new Span(area.LeftColumn, area.RightColumn, area.TopRow, area.BottomRow);
            }

            internal bool CrossIsInside(int first, int last) => CrossFirst >= first && CrossLast <= last;

            internal Area ToArea(bool rows)
            {
                return rows
                    ? new Area(FirstLine, CrossFirst, LastLine, CrossLast)
                    : new Area(CrossFirst, FirstLine, CrossLast, LastLine);
            }
        }

        /// <summary>
        /// The steps of an operation that does many inserts, deletes or other changes. Each step changes
        /// the ranges of the array formulas and data tables in the same way as the operation changes the
        /// sheet. A step throws when the operation changes only part of a range.
        /// </summary>
        internal sealed class OperationCheck
        {
            private readonly XLWorksheet _sheet;
            private readonly List<Step> _steps = new();

            internal OperationCheck(XLWorksheet sheet)
            {
                _sheet = sheet;
            }

            /// <summary>
            /// Add an insert of cells.
            /// </summary>
            /// <param name="rows"><c>true</c> when the insert moves cells down, <c>false</c> when it moves cells right.</param>
            /// <param name="firstLine">The first line that moves.</param>
            /// <param name="numberOfLines">The number of inserted lines.</param>
            /// <param name="crossFirst">The first line of the cross span that moves.</param>
            /// <param name="crossLast">The last line of the cross span that moves.</param>
            /// <returns>This check.</returns>
            internal OperationCheck Insert(bool rows, int firstLine, int numberOfLines, int crossFirst, int crossLast)
            {
                _steps.Add(new Step(
                    LinesToEnd(rows, firstLine, crossFirst, crossLast),
                    ranges => ApplyInsert(ranges, rows, firstLine, numberOfLines, crossFirst, crossLast)));
                return this;
            }

            /// <summary>
            /// Add a delete of cells.
            /// </summary>
            /// <param name="rows"><c>true</c> when the delete moves cells up, <c>false</c> when it moves cells left.</param>
            /// <param name="deleted">The deleted cells.</param>
            /// <returns>This check.</returns>
            internal OperationCheck Delete(bool rows, Area deleted)
            {
                var span = Span.Of(deleted, rows);
                _steps.Add(new Step(
                    LinesToEnd(rows, span.FirstLine, span.CrossFirst, span.CrossLast),
                    ranges => ApplyDelete(ranges, rows, span)));
                return this;
            }

            /// <summary>
            /// Add a clear of cells. A clear removes an array formula that is fully in the cleared cells.
            /// </summary>
            /// <param name="cleared">The cleared cells.</param>
            /// <returns>This check.</returns>
            internal OperationCheck Clear(Area cleared)
            {
                _steps.Add(new Step(cleared, ranges => ApplyClear(ranges, cleared)));
                return this;
            }

            /// <summary>
            /// Add a transpose of the cells in a square. The cells on the diagonal stay, and the other
            /// cells move. The step cannot change the range of an array formula, so the cells that move
            /// must not have one.
            /// </summary>
            /// <param name="square">The square with the transposed cells.</param>
            /// <returns>This check.</returns>
            internal OperationCheck Transpose(Area square)
            {
                _steps.Add(new Step(square, ranges => ApplyTranspose(ranges, square)));
                return this;
            }

            /// <summary>
            /// Check all steps. Then suspend the check of each step until you dispose the returned object.
            /// </summary>
            /// <remarks>
            /// The check reads the formulas that the search of each step reads. A step moves a range only
            /// inside the area of the step, so a range outside the areas of all steps stays.
            /// </remarks>
            /// <returns>An object that resumes the check when you dispose it.</returns>
            /// <exception cref="InvalidOperationException">A step changes only part of an array formula,
            /// or the transpose moves cells of an array formula.</exception>
            internal IDisposable Begin()
            {
                if (_steps.Count > 0 && _sheet.ArrayFormulaChecksSuspended == 0)
                {
                    var ranges = GetMovableRanges(_sheet, _steps.Select(static step => step.Area).ToList());
                    foreach (var step in _steps)
                        step.Apply(ranges);

                    ThrowIfPartlyDeleted(ranges);
                }

                return SuspendChecks(_sheet);
            }

            private static void ApplyClear(List<MovableRange> ranges, Area cleared)
            {
                for (var i = ranges.Count - 1; i >= 0; --i)
                {
                    var current = ranges[i].Current;
                    var intersection = current.Intersect(cleared);
                    if (intersection is null)
                        continue;

                    if (intersection.Value != current)
                        throw CreateException(ranges[i].Range);

                    ranges.RemoveAt(i);
                }
            }

            private static void ApplyTranspose(List<MovableRange> ranges, Area square)
            {
                foreach (var range in ranges)
                {
                    // Two or more cells of the square always have a cell that is not on the diagonal.
                    if (range.Current.Intersect(square) is { } intersection && !IsOneCellOnDiagonal(intersection, square))
                        throw CreateTransposeException(range.Range);
                }
            }

            private static bool IsOneCellOnDiagonal(Area cells, Area square)
            {
                return cells.TopRow == cells.BottomRow
                    && cells.LeftColumn == cells.RightColumn
                    && cells.TopRow - square.TopRow == cells.LeftColumn - square.LeftColumn;
            }

            /// <summary>
            /// A step of the operation.
            /// </summary>
            /// <param name="Area">The cells that the step can change. The step moves a range only inside this area.</param>
            /// <param name="Apply">Change the ranges in the same way as the step changes the sheet.</param>
            private sealed record Step(Area Area, Action<List<MovableRange>> Apply);
        }

        /// <summary>
        /// An array formula or a data table during a check of many steps.
        /// </summary>
        /// <param name="range">The range before the steps.</param>
        private sealed class MovableRange(Area range)
        {
            internal Area Range { get; } = range;

            /// <summary>
            /// The range after the steps that the check applied.
            /// </summary>
            internal Area Current { get; set; } = range;

            /// <summary>
            /// <c>true</c> when a delete removed some lines of the range.
            /// </summary>
            internal bool LostLines { get; set; }
        }

        private sealed class CheckResumers(List<IDisposable> resumers) : IDisposable
        {
            public void Dispose()
            {
                resumers.ForEach(static resumer => resumer.Dispose());
                resumers.Clear();
            }
        }

        private sealed class CheckResumer : IDisposable
        {
            private XLWorksheet? _sheet;

            internal CheckResumer(XLWorksheet sheet) => _sheet = sheet;

            public void Dispose()
            {
                if (_sheet is null)
                    return;

                _sheet.ArrayFormulaChecksSuspended--;
                _sheet = null;
            }
        }
    }
}
