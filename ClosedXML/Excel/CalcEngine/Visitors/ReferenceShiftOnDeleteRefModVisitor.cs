using ClosedXML.Extensions;
using ClosedXML.Parser;
using System;

namespace ClosedXML.Excel.CalcEngine.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is deleted and shifted up/left.
/// </summary>
internal class ReferenceShiftOnDeleteRefModVisitor : CopyVisitor
{
    private readonly SheetArea _deletedBookArea;
    private readonly XLShiftDeletedCells _shift;

    public ReferenceShiftOnDeleteRefModVisitor(SheetArea deletedBookArea, XLShiftDeletedCells shift)
    {
        _deletedBookArea = deletedBookArea;
        _shift = shift;
    }

    public override TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, ReadOnlySpan<char> sheet, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, sheet, reference);
    }

    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, ReadOnlySpan<char>.Empty, reference);
    }

    // The referenceSheetName is empty, if the reference has no sheet.
    private TransformedSymbol ShiftFormulaReferences(ModContext ctx, SymbolRange range, ReadOnlySpan<char> referenceSheetName, ReferenceArea referenceToShift)
    {
        var sheetName = referenceSheetName.IsEmpty ? ctx.Sheet.AsSpan() : referenceSheetName;
        if (!XLHelper.SheetNameEquals(sheetName, _deletedBookArea.Name))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var wouldSplitArea = _shift == XLShiftDeletedCells.ShiftCellsUp
            ? !referenceToShift.TryDeleteAndShiftUp(_deletedBookArea.Area, out var shiftedReference)
            : !referenceToShift.TryDeleteAndShiftLeft(_deletedBookArea.Area, out shiftedReference);

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
