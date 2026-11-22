using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// The options of a data bar.
    /// </summary>
    public interface IXLCFDataBar
    {
        /// <inheritdoc cref="IXLConditionalFormat.Gradient"/>
        Boolean Gradient { get; }

        /// <summary>
        /// Set the fill of the data bar.
        /// </summary>
        /// <param name="value"><c>true</c> for a gradient fill, <c>false</c> for a solid fill.</param>
        IXLCFDataBar SetGradient(Boolean value = true);
    }
}
