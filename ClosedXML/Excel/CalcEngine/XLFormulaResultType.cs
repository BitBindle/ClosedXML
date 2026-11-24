namespace ClosedXML.Excel
{
    /// <summary>
    /// The type of an <see cref="XLFormulaResult"/>.
    /// </summary>
    public enum XLFormulaResultType
    {
        /// <summary>
        /// A single value: blank, a logical value, a number, a text or an error.
        /// </summary>
        Value,

        /// <summary>
        /// A reference to one or more ranges.
        /// </summary>
        Reference,

        /// <summary>
        /// An array of values.
        /// </summary>
        Array
    }
}
