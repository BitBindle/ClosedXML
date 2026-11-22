using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeColumns : IEnumerable<IXLRangeColumn>
    {
        /// <summary>
        /// Adds a column range to this group.
        /// </summary>
        /// <param name="columRange">The column range to add.</param>
        void Add(IXLRangeColumn columRange);

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
        /// <param name="options">The options to determine whether a cell is used.</param>
        IXLCells CellsUsed(XLCellsUsedOptions options);

        /// <summary>
        /// Deletes all columns and shifts the columns at the right of them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjust the width of the worksheet columns to the contents of the cells in these range columns.
        /// The <c>startRow</c> and <c>endRow</c> parameters are row numbers of the worksheet.
        /// A missing start or end row is the first or last row of the range.
        /// </summary>
        IXLRangeColumns AdjustToContents();

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeColumns AdjustToContents(Int32 startRow);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeColumns AdjustToContents(Int32 startRow, Int32 endRow);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeColumns AdjustToContents(Double minWidth, Double maxWidth);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeColumns AdjustToContents(Int32 startRow, Double minWidth, Double maxWidth);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeColumns AdjustToContents(Int32 startRow, Int32 endRow, Double minWidth, Double maxWidth);

        /// <summary>
        /// Get a new collection with the columns that match the <paramref name="predicate"/>.
        /// The columns keep the order of the enumeration.
        /// </summary>
        /// <param name="predicate">The condition for a column.</param>
        /// <exception cref="ArgumentNullException">The <paramref name="predicate"/> is null.</exception>
        IXLRangeColumns Where(Func<IXLRangeColumn, Boolean> predicate);

        /// <summary>
        /// Get a new collection without the first <paramref name="count"/> columns of the enumeration.
        /// </summary>
        /// <param name="count">The number of columns to remove. A value of 0 or less removes no column.</param>
        IXLRangeColumns Skip(Int32 count);

        /// <summary>
        /// Get a new collection with the first <paramref name="count"/> columns of the enumeration.
        /// </summary>
        /// <param name="count">The number of columns to keep. A value of 0 or less keeps no column.</param>
        IXLRangeColumns Take(Int32 count);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these columns.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
