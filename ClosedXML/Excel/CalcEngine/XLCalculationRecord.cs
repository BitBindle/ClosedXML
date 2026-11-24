namespace ClosedXML.Excel
{
    /// <summary>
    /// The kind of an operation that a formula used during a recalculation.
    /// </summary>
    /// <seealso cref="IXLWorkbook.RecalculateAllFormulas(System.Collections.Generic.ICollection{IXLCell}, System.Collections.Generic.ICollection{XLCalculationRecord})"/>
    public enum XLCalculationRecordKind
    {
        /// <summary>
        /// The result of the operation can depend on the language or the locale of the application. Examples are
        /// a conversion of a text with a digit to a number, a conversion of a number to a text, an order comparison
        /// of two texts and the upper case of a text with the letter <c>i</c>.
        /// </summary>
        LocaleOperation,

        /// <summary>
        /// The calculation engine can give another result than Excel for the operation. Examples are a criterion
        /// longer than 255 characters and a wildcard pattern with a <c>~</c> that is not followed by <c>*</c>,
        /// <c>?</c> or <c>~</c>.
        /// </summary>
        EngineDifference,
    }

    /// <summary>
    /// A formula cell that used an operation of a kind during a recalculation.
    /// </summary>
    /// <param name="Cell">The formula cell.</param>
    /// <param name="Kind">The kind of the operation.</param>
    public readonly record struct XLCalculationRecord(IXLCell Cell, XLCalculationRecordKind Kind);
}
