using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// The direction of the bars of a data bar.
    /// </summary>
    internal enum XLCFDataBarDirection
    {
        /// <summary>The reading order of the cell gives the direction.</summary>
        Context,
        LeftToRight,
        RightToLeft,
    }

    /// <summary>
    /// The position of the axis of a data bar with negative values.
    /// </summary>
    internal enum XLCFDataBarAxisPosition
    {
        /// <summary>The negative and positive values give the position of the axis.</summary>
        Automatic,

        /// <summary>The axis is at the middle of the cell.</summary>
        Middle,

        /// <summary>The data bar has no axis. Negative bars go in the direction of positive bars.</summary>
        None,
    }

    /// <summary>
    /// The options of an Excel 2010 data bar (the <c>x14:dataBar</c> element) that the public API does not set.
    /// The load reads them from a file, and the save writes them again.
    /// </summary>
    /// <remarks>
    /// The initial values are the values that the save writes for a new data bar.
    /// </remarks>
    internal sealed record XLCFDataBarOptions
    {
        internal static XLCFDataBarOptions Default { get; } = new();

        /// <summary>The length of the shortest bar, in percent of the cell width.</summary>
        public UInt32 MinLength { get; init; }

        /// <summary>The length of the longest bar, in percent of the cell width.</summary>
        public UInt32 MaxLength { get; init; } = 100;

        public Boolean Border { get; init; }

        /// <summary>The border color of the positive bars. A null value writes no color.</summary>
        public XLColor? BorderColor { get; init; }

        public Boolean NegativeBarColorSameAsPositive { get; init; }

        public Boolean NegativeBarBorderColorSameAsPositive { get; init; } = true;

        /// <summary>The border color of the negative bars. A null value writes no color.</summary>
        public XLColor? NegativeBorderColor { get; init; }

        public XLCFDataBarDirection Direction { get; init; }

        public XLCFDataBarAxisPosition AxisPosition { get; init; }

        /// <summary>The color of the axis. A null value writes no color.</summary>
        public XLColor? AxisColor { get; init; } = XLColor.Black;
    }
}
