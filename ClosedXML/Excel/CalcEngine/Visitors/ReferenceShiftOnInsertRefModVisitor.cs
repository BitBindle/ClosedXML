using ClosedXML.Extensions;
using ClosedXML.Parser;
using System;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is inserted and cells are shifted down/right.
/// </summary>
internal class ReferenceShiftOnInsertRefModVisitor : CopyVisitor
{
    private readonly SheetArea _insertedBookArea;
    private readonly bool _shiftDown;

    internal ReferenceShiftOnInsertRefModVisitor(SheetArea insertedBookArea, bool shiftDown)
    {
        _insertedBookArea = insertedBookArea;
        _shiftDown = shiftDown;
    }

    public override TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, ReadOnlySpan<char> sheet, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, sheet, reference);
    }

    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, ReadOnlySpan<char>.Empty, reference);
    }

    public override TransformedSymbol Unary(ModContext ctx, SymbolRange range, UnaryOperation operation, TransformedSymbol arg)
    {
        // The spill of a removed reference is the error, not the error and a spill operator.
        if (operation == UnaryOperation.SpillRange && arg.AsSpan().SequenceEqual(XLHelper.RefError))
            return TransformedSymbol.ToText(ctx.Formula, range, XLHelper.RefError);

        return base.Unary(ctx, range, operation, arg);
    }

    // The referenceSheetName is empty, if the reference has no sheet.
    private TransformedSymbol ShiftFormulaReferences(ModContext ctx, SymbolRange range, ReadOnlySpan<char> referenceSheetName, ReferenceArea referenceToShift)
    {
        var sheetName = referenceSheetName.IsEmpty ? ctx.Sheet.AsSpan() : referenceSheetName;
        if (!XLHelper.SheetNameEquals(sheetName, _insertedBookArea.Name))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var wouldSplitArea = _shiftDown
            ? !referenceToShift.TryInsertAndShiftDown(_insertedBookArea.Area, out var shiftedReference)
            : !referenceToShift.TryInsertAndShiftRight(_insertedBookArea.Area, out shiftedReference);

        // Return original reference if the shift would cause a split
        if (wouldSplitArea)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // If reference was shifted out of sheet, return #REF!
        if (shiftedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, XLHelper.RefError);

        // Do not allocate a new string unless necessary
        if (referenceToShift == shiftedReference.Value)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var shiftedReferenceA1 = shiftedReference.Value.GetDisplayStringA1(referenceSheetName.IsEmpty ? null : referenceSheetName.ToString());
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReferenceA1);
    }
}
