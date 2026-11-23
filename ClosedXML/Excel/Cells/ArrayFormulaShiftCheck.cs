using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    /// <summary>
    /// Stops an insert or a delete that moves only part of an array formula or a data table.
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

            var maxLine = rows ? XLHelper.MaxRowNumber : XLHelper.MaxColumnNumber;
            foreach (var range in GetArrayRanges(sheet, rows, firstLine, crossFirst, crossLast))
            {
                // The range moves as one rectangle only when it is fully in the cells that move. The insert
                // removes the cells that move out of the sheet, so the range must also stay in the sheet.
                var span = Span.Of(range, rows);
                if (span.FirstLine >= firstLine && span.CrossIsInside(crossFirst, crossLast) && (long)span.LastLine + numberOfLines <= maxLine)
                    continue;

                throw CreateException(range);
            }
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
            ApplyDelete(ranges, span);
            ThrowIfPartlyDeleted(ranges);
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
                    ApplyDelete(ranges, part);

                ThrowIfPartlyDeleted(ranges);
                return;
            }

            // The sheet moves and changes the other range objects after each delete. This check does not repeat that,
            // so it does not allow an array formula in the cells that can change. A part can change only the cells in
            // its cross span, from its first line to the end of the sheet.
            var changedRange = ranges.Find(range => parts.Any(part => CanChange(part, range.Span)));
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
        private static void ApplyDelete(List<MovableRange> ranges, Span deleted)
        {
            var deletedCount = deleted.LastLine - deleted.FirstLine + 1;
            for (var i = ranges.Count - 1; i >= 0; --i)
            {
                var range = ranges[i];
                var span = range.Span;
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
                range.Span = span with { FirstLine = newFirst, LastLine = newLast };
                range.LostLines |= lostCount > 0;
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
            return GetArrayRanges(sheet, rows, firstLine, crossFirst, crossLast)
                .Select(range => new MovableRange(range, Span.Of(range, rows)))
                .ToList();
        }

        /// <summary>
        /// Get the ranges of the array formulas and data tables in the cross span. Each range has cells
        /// in the lines from <paramref name="firstLine"/> to the end of the sheet. Only these ranges can
        /// move or lose cells.
        /// </summary>
        /// <remarks>
        /// All cells of an array formula share one formula. A data table from a file has the formula only
        /// in its first cell. This cell can be before the area, so the search starts at the first row and
        /// the first column of the sheet.
        /// </remarks>
        private static List<Area> GetArrayRanges(XLWorksheet sheet, bool rows, int firstLine, int crossFirst, int crossLast)
        {
            var area = rows
                ? new Area(firstLine, crossFirst, XLHelper.MaxRowNumber, crossLast)
                : new Area(crossFirst, firstLine, crossLast, XLHelper.MaxColumnNumber);
            var searchArea = new Area(1, 1, area.BottomRow, area.RightColumn);
            var ranges = new List<Area>();
            var foundFormulas = new HashSet<XLCellFormula>(ReferenceEqualityComparer.Instance);
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(searchArea);
            while (enumerator.MoveNext())
            {
                var formula = enumerator.Current;
                if (formula is null || formula.Type is not (FormulaType.Array or FormulaType.DataTable) || !formula.Range.Intersects(area))
                    continue;

                if (foundFormulas.Add(formula))
                    ranges.Add(formula.Range);
            }

            return ranges;
        }

        private static InvalidOperationException CreateException(Area range)
        {
            return new InvalidOperationException($"The operation changes only part of the array formula or data table in {range}. Change all cells of the range, or remove the formula first.");
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
        }

        /// <summary>
        /// An array formula or a data table during a check of many deletes.
        /// </summary>
        /// <param name="range">The range before the deletes.</param>
        /// <param name="span">The lines of the range after the deletes that the check applied.</param>
        private sealed class MovableRange(Area range, Span span)
        {
            internal Area Range { get; } = range;

            internal Span Span { get; set; } = span;

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
