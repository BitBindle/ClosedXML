using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine;

/// <summary>
/// Collects the formula cells that use an operation of a <see cref="XLCalculationRecordKind"/> during one
/// recalculation. Each cell gets at most one record of each kind.
/// </summary>
/// <remarks>
/// <para>
/// The conversions of <see cref="ScalarValue"/>, <see cref="Wildcard"/> and <see cref="ScalarValueComparer"/>
/// have no <see cref="CalcContext"/>, and other code calls them too. So the recorder is a value of the thread,
/// not of the context. Each recalculation sets its recorder, or <c>null</c> when it has no collection of records,
/// and restores the previous value at the end. Without a recorder, a hook does nothing.
/// </para>
/// <para>
/// A cell can stop its evaluation to wait for a dirty cell, and the engine evaluates the cell again later.
/// So the recorder keeps the kinds of the current cell, and adds the records only when the cell is complete.
/// </para>
/// </remarks>
internal sealed class CalculationRecorder
{
    [ThreadStatic]
    private static CalculationRecorder? t_current;

    private static readonly XLCalculationRecordKind[] Kinds =
    {
        XLCalculationRecordKind.LocaleOperation,
        XLCalculationRecordKind.EngineDifference,
    };

    private readonly ICollection<XLCalculationRecord> _records;
    private XLWorksheet? _sheet;
    private Point _point;

    /// <summary>The kinds that the current cell used, as bits.</summary>
    private int _kinds;

    internal CalculationRecorder(ICollection<XLCalculationRecord> records)
    {
        _records = records;
    }

    /// <summary>
    /// The recorder of the recalculation on this thread, or <c>null</c> when the recalculation doesn't collect records.
    /// </summary>
    internal static CalculationRecorder? Current
    {
        get => t_current;
        set => t_current = value;
    }

    /// <summary>
    /// Gets the local time of a clock of the caller. The clock can run code of the caller, e.g. an evaluation of
    /// another workbook. So the clock runs without a recorder, and its code can't report into the current cell.
    /// </summary>
    internal static DateTimeOffset GetLocalNow(TimeProvider timeProvider)
    {
        var recorder = t_current;
        if (recorder is null)
            return timeProvider.GetLocalNow();

        t_current = null;
        try
        {
            return timeProvider.GetLocalNow();
        }
        finally
        {
            t_current = recorder;
        }
    }

    /// <summary>
    /// Calls a name filter of the caller without a recorder, for the same reason as <see cref="GetLocalNow"/>.
    /// </summary>
    internal static bool CallFilter(Func<string?, string, bool> filter, string? sheetName, string name)
    {
        var recorder = t_current;
        if (recorder is null)
            return filter(sheetName, name);

        t_current = null;
        try
        {
            return filter(sheetName, name);
        }
        finally
        {
            t_current = recorder;
        }
    }

    /// <summary>
    /// Starts the operations of a formula cell.
    /// </summary>
    internal void BeginCell(XLWorksheet sheet, Point point)
    {
        _sheet = sheet;
        _point = point;
        _kinds = 0;
    }

    /// <summary>
    /// Ends the current cell and returns its kinds. Then no cell is current, so code that runs before
    /// <see cref="Write"/> (e.g. the collection of failed cells) can't report into the cell.
    /// </summary>
    internal CellKinds TakeCell()
    {
        var kinds = new CellKinds(_sheet, _point, _kinds);
        DiscardCell();
        return kinds;
    }

    /// <summary>
    /// Adds a record for each kind of a cell from <see cref="TakeCell"/>.
    /// </summary>
    internal void Write(CellKinds cell)
    {
        if (cell.Sheet is null || cell.Kinds == 0)
            return;

        var xlCell = cell.Sheet.Cell(cell.Point.Row, cell.Point.Column);
        foreach (var kind in Kinds)
        {
            if ((cell.Kinds & Bit(kind)) != 0)
                _records.Add(new XLCalculationRecord(xlCell, kind));
        }
    }

    /// <summary>
    /// Forgets the operations of the current cell, because the engine evaluates the cell again later.
    /// </summary>
    internal void DiscardCell()
    {
        _sheet = null;
        _kinds = 0;
    }

    /// <summary>Records an operation whose result can depend on the locale.</summary>
    internal void Locale() => Report(XLCalculationRecordKind.LocaleOperation);

    /// <summary>Records an operation for which the engine can give another result than Excel.</summary>
    internal void EngineDifference() => Report(XLCalculationRecordKind.EngineDifference);

    /// <summary>
    /// Records a locale operation when the text has a digit. Each locale reads a text without a digit as no number
    /// and no date.
    /// </summary>
    internal void LocaleIfDigit(ReadOnlySpan<char> text)
    {
        foreach (var c in text)
        {
            if (char.IsDigit(c))
            {
                Locale();
                return;
            }
        }
    }

    /// <summary>
    /// Records a locale operation for the upper or the lower case of a text with the letter i in one of its forms
    /// <c>i</c>, <c>I</c>, <c>ı</c> or <c>İ</c>. The case of these letters differs in Turkish and Azerbaijani.
    /// </summary>
    internal void LocaleIfLetterI(ReadOnlySpan<char> text)
    {
        if (HasLetterI(text))
            Locale();
    }

    /// <summary>
    /// Records a locale operation for an equality of two texts when both texts have the letter i in one of its forms.
    /// Only these letters are equal without case in one language and not equal in another.
    /// </summary>
    internal void LocaleIfBothHaveLetterI(ReadOnlySpan<char> first, ReadOnlySpan<char> second)
    {
        if (HasLetterI(first) && HasLetterI(second))
            Locale();
    }

    private static bool HasLetterI(ReadOnlySpan<char> text) => text.IndexOfAny("iIıİ") >= 0;

    private void Report(XLCalculationRecordKind kind)
    {
        // Between two cells, no cell is current.
        if (_sheet is not null)
            _kinds |= Bit(kind);
    }

    private static int Bit(XLCalculationRecordKind kind) => 1 << (int)kind;

    /// <summary>
    /// The kinds of a formula cell that has ended.
    /// </summary>
    internal readonly record struct CellKinds(XLWorksheet? Sheet, Point Point, int Kinds);
}
