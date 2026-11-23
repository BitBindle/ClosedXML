using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLRangeRows : IEnumerable<IXLRangeRow>
    {
        /// <summary>
        /// Adds a row range to this group.
        /// </summary>
        /// <param name="rowRange">The row range to add.</param>
        void Add(IXLRangeRow rowRange);

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
        /// Deletes all rows and shifts the rows below them accordingly.
        /// </summary>
        void Delete();

        /// <summary>
        /// Adjust the height of the worksheet rows to the contents of the cells in these range rows.
        /// The <c>startColumn</c> and <c>endColumn</c> parameters are column numbers of the worksheet.
        /// A missing start or end column is the first or last column of the range.
        /// </summary>
        IXLRangeRows AdjustToContents();

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeRows AdjustToContents(Int32 startColumn);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeRows AdjustToContents(Int32 startColumn, Int32 endColumn);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeRows AdjustToContents(Double minHeight, Double maxHeight);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight);

        /// <inheritdoc cref="AdjustToContents()"/>
        IXLRangeRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight);

        /// <summary>
        /// Get a new collection with the rows that match the <paramref name="predicate"/>.
        /// The rows keep the order of the enumeration.
        /// </summary>
        /// <param name="predicate">The condition for a row.</param>
        /// <exception cref="ArgumentNullException">The <paramref name="predicate"/> is null.</exception>
        IXLRangeRows Where(Func<IXLRangeRow, Boolean> predicate);

        /// <summary>
        /// Get a new collection without the first <paramref name="count"/> rows of the enumeration.
        /// </summary>
        /// <param name="count">The number of rows to remove. A value of 0 or less removes no row.</param>
        IXLRangeRows Skip(Int32 count);

        /// <summary>
        /// Get a new collection with the first <paramref name="count"/> rows of the enumeration.
        /// </summary>
        /// <param name="count">The number of rows to keep. A value of 0 or less keeps no row.</param>
        IXLRangeRows Take(Int32 count);

        IXLStyle Style { get; set; }

        /// <summary>
        /// Clears the contents of these rows.
        /// </summary>
        /// <param name="clearOptions">Specify what you want to clear.</param>
        /// <exception cref="InvalidOperationException">The clear removes the contents of only part of an array formula or a data table.</exception>
        IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

        void Select();
    }
}
