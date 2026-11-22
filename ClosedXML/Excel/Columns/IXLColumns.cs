using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLColumns : IEnumerable<IXLColumn>
    {
        /// <summary>
        /// Sets the width of all columns.
        /// </summary>
        /// <value>
        /// The width of all columns.
        /// </value>
        Double Width { set; }

        /// <summary>
        /// Deletes all columns and shifts the columns at the right of them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjusts the width of all columns based on its contents.
        /// </summary>
        IXLColumns AdjustToContents();

        /// <summary>
        /// Adjusts the width of all columns based on its contents, starting from the startRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        IXLColumns AdjustToContents(Int32 startRow);

        /// <summary>
        /// Adjusts the width of all columns based on its contents, starting from the startRow and ending at endRow.
        /// </summary>
        /// <param name="startRow">The row to start calculating the column width.</param>
        /// <param name="endRow">The row to end calculating the column width.</param>
        IXLColumns AdjustToContents(Int32 startRow, Int32 endRow);

        IXLColumns AdjustToContents(Double minWidth, Double maxWidth);

        IXLColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth);

        IXLColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth);

        /// <summary>
        /// Get a new collection with the columns that match the <paramref name="predicate"/>.
        /// The columns keep the order of the enumeration.
        /// </summary>
        /// <param name="predicate">The condition for a column.</param>
        /// <exception cref="ArgumentNullException">The <paramref name="predicate"/> is null.</exception>
        IXLColumns Where(Func<IXLColumn, Boolean> predicate);

        /// <summary>
        /// Get a new collection without the first <paramref name="count"/> columns of the enumeration.
        /// </summary>
        /// <param name="count">The number of columns to remove. A value of 0 or less removes no column.</param>
        IXLColumns Skip(Int32 count);

        /// <summary>
        /// Get a new collection with the first <paramref name="count"/> columns of the enumeration.
        /// </summary>
        /// <param name="count">The number of columns to keep. A value of 0 or less keeps no column.</param>
        IXLColumns Take(Int32 count);

        /// <summary>
        /// Hides all columns.
        /// </summary>
        void Hide();

        /// <summary>Unhides all columns.</summary>
        void Unhide();

        /// <summary>
        /// Increments the outline level of all columns by 1.
        /// </summary>
        void Group();

        /// <summary>
        /// Increments the outline level of all columns by 1.
        /// </summary>
        /// <param name="collapse">If set to <c>true</c> the columns will be shown collapsed.</param>
        void Group(Boolean collapse);

        /// <summary>
        /// Sets outline level for all columns.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        void Group(Int32 outlineLevel);

        /// <summary>
        /// Sets outline level for all columns.
        /// </summary>
        /// <param name="outlineLevel">The outline level.</param>
        /// <param name="collapse">If set to <c>true</c> the columns will be shown collapsed.</param>
        void Group(Int32 outlineLevel, Boolean collapse);

        /// <summary>
        /// Decrements the outline level of all columns by 1.
        /// </summary>
        void Ungroup();

        /// <summary>
        /// Decrements the outline level of all columns by 1.
        /// </summary>
        /// <param name="fromAll">If set to <c>true</c> it will remove the columns from all outline levels.</param>
        void Ungroup(Boolean fromAll);

        /// <summary>
        /// Show all columns as collapsed.
        /// </summary>
        void Collapse();

        /// <summary>Expands all columns (if they're collapsed).</summary>
        void Expand();

        /// <summary>
        /// Returns the collection of cells.
        /// </summary>
        IXLCells Cells();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        IXLCells CellsUsed();

        /// <summary>
        /// Returns the collection of cells that have a value.
        /// </summary>
        /// <param name="includeFormats">if set to <c>true</c> will return all cells with a value or a style different than the default.</param>
        IXLCells CellsUsed(Boolean includeFormats);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Adds a vertical page break after these columns.
        /// </summary>
        IXLColumns AddVerticalPageBreaks();

        /// <summary>
        /// Clears the contents of these columns.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLColumns Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
