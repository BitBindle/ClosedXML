// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using ClosedXML.Excel.CalcEngine.Exceptions;

namespace ClosedXML.Excel
{
    public interface IXLWorkbook : IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>, IDisposable
    {
        String Author { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's calculation mode.
        /// </summary>
        XLCalculateMode CalculateMode { get; set; }

        Boolean CalculationOnSave { get; set; }

        /// <summary>
        ///   Gets or sets the default column width for the workbook.
        ///   <para>All new worksheets will use this column width.</para>
        /// </summary>
        Double ColumnWidth { get; set; }

        IXLCustomProperties CustomProperties { get; }

        Boolean DefaultRightToLeft { get; }

        Boolean DefaultShowFormulas { get; }

        Boolean DefaultShowGridLines { get; }

        Boolean DefaultShowOutlineSymbols { get; }

        Boolean DefaultShowRowColHeaders { get; }

        Boolean DefaultShowRuler { get; }

        Boolean DefaultShowWhiteSpace { get; }

        Boolean DefaultShowZeros { get; }

        IXLFileSharing FileSharing { get; }

        Boolean ForceFullCalculation { get; set; }

        Boolean FullCalculationOnLoad { get; set; }

        Boolean FullPrecision { get; set; }

        Boolean LockStructure { get; set; }

        Boolean LockWindows { get; set; }

        [Obsolete($"Use {nameof(DefinedNames)} instead.")]
        IXLDefinedNames NamedRanges { get; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's defined names.
        /// </summary>
        IXLDefinedNames DefinedNames { get; }

        /// <summary>
        ///   Gets or sets the default outline options for the workbook.
        ///   <para>All new worksheets will use these outline options.</para>
        /// </summary>
        IXLOutline Outline { get; set; }

        /// <summary>
        ///   Gets or sets the default page options for the workbook.
        ///   <para>All new worksheets will use these page options.</para>
        /// </summary>
        IXLPageSetup PageOptions { get; set; }

        /// <summary>
        ///   Gets all pivot caches in a workbook. A one cache can be
        ///   used by multiple tables. Unused caches are not saved.
        /// </summary>
        IXLPivotCaches PivotCaches { get; }

        /// <summary>
        ///   Gets or sets the workbook's properties.
        /// </summary>
        XLWorkbookProperties Properties { get; set; }

        /// <summary>
        ///   Gets or sets the workbook's reference style.
        /// </summary>
        XLReferenceStyle ReferenceStyle { get; set; }

        Boolean RightToLeft { get; set; }

        /// <summary>
        ///   Gets or sets the default row height for the workbook.
        ///   <para>All new worksheets will use this row height.</para>
        /// </summary>
        Double RowHeight { get; set; }

        Boolean ShowFormulas { get; set; }

        Boolean ShowGridLines { get; set; }

        Boolean ShowOutlineSymbols { get; set; }

        Boolean ShowRowColHeaders { get; set; }

        Boolean ShowRuler { get; set; }

        Boolean ShowWhiteSpace { get; set; }

        Boolean ShowZeros { get; set; }

        /// <summary>
        /// Gets or sets the default format of the workbook. All cells that don't have specified format, either
        /// at sheet, column, row or cell level, will use this format.
        /// </summary>
        IXLStyle Style { get; set; }

        /// <summary>
        ///   Gets an object to manipulate this workbook's theme.
        /// </summary>
        IXLTheme Theme { get; }

        Boolean Use1904DateSystem { get; set; }

        /// <summary>
        ///   Gets an object to manipulate the worksheets.
        /// </summary>
        IXLWorksheets Worksheets { get; }

        IXLWorksheet AddWorksheet();

        IXLWorksheet AddWorksheet(Int32 position);

        IXLWorksheet AddWorksheet(String sheetName);

        IXLWorksheet AddWorksheet(String sheetName, Int32 position);

        void AddWorksheet(DataSet dataSet);

        void AddWorksheet(IXLWorksheet worksheet);

        /// <summary>
        /// Add a worksheet with a table at Cell(row:1, column:1). The dataTable's name is used for the
        /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
        /// </summary>
        /// <param name="dataTable">Datatable to insert</param>
        /// <returns>Inserted Worksheet</returns>
        IXLWorksheet AddWorksheet(DataTable dataTable);

        /// <summary>
        /// Add a worksheet with a table at Cell(row:1, column:1). The sheetName provided is used for the
        /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
        /// </summary>
        /// <param name="dataTable">dataTable to insert as Excel Table</param>
        /// <param name="sheetName">Worksheet and Excel Table name</param>
        /// <returns>Inserted Worksheet</returns>
        IXLWorksheet AddWorksheet(DataTable dataTable, String sheetName);

        /// <summary>
        /// Add a worksheet with a table at Cell(row:1, column:1).
        /// </summary>
        /// <param name="dataTable">dataTable to insert as Excel Table</param>
        /// <param name="sheetName">Worksheet name</param>
        /// <param name="tableName">Excel Table name</param>
        /// <returns>Inserted Worksheet</returns>
        IXLWorksheet AddWorksheet(DataTable dataTable, String sheetName, String tableName);

        /// <summary>
        /// Get a cell from a defined name or from a full address with a sheet (e.g. <c>Sheet!A1</c>).
        /// </summary>
        /// <returns>The cell, or null when the name or the address does not refer to a cell.</returns>
        IXLCell? Cell(String namedCell);

        IXLCells Cells(String namedCells);

        IXLCustomProperty CustomProperty(String name);

        /// <summary>
        /// Evaluate a formula expression.
        /// </summary>
        /// <param name="expression">Formula expression to evaluate.</param>
        /// <exception cref="MissingContextException">
        /// If the expression contains a function that requires a context (e.g. current cell or worksheet).
        /// </exception>
        XLCellValue Evaluate(String expression);

        IXLCells FindCells(Func<IXLCell, Boolean> predicate);

        IXLColumns FindColumns(Func<IXLColumn, Boolean> predicate);

        IXLRows FindRows(Func<IXLRow, Boolean> predicate);

        [Obsolete($"Use {nameof(DefinedName)} instead.")]
        IXLDefinedName? NamedRange(String name);

        /// <summary>
        /// Try to find a defined name. If <paramref name="name"/> specifies a sheet, try to find
        /// name in the sheet first and fall back to the workbook if not found in the sheet.
        /// <para>
        /// <example>
        /// Requested name <c>Sheet1!Name</c> will first try to find <c>Name</c> in a sheet
        /// <c>Sheet1</c> (if such sheet exists) and if not found there, tries to find <c>Name</c>
        /// in workbook.
        /// </example>
        /// </para>
        /// <para>
        /// <example>
        /// Requested name <c>Name</c> will be searched only in a workbooks <see cref="DefinedNames"/>.
        /// </example>
        /// </para>
        /// </summary>
        /// <param name="name">Name of requested name, either plain name (e.g. <c>Name</c>) or with
        /// sheet specified (e.g. <c>Sheet!Name</c>).</param>
        /// <returns>Found name or null.</returns>
        IXLDefinedName? DefinedName(String name);

        /// <summary>
        /// Get a range from a defined name or from a full address with a sheet (e.g. <c>Sheet!A1:B2</c>).
        /// </summary>
        /// <returns>The range, or null when the name or the address does not refer to a range.</returns>
        IXLRange? Range(String range);

        /// <summary>
        /// Get a range from a full address with a sheet (e.g. <c>Sheet!A1:B2</c>).
        /// </summary>
        /// <param name="rangeAddress">The address of the range with a sheet.</param>
        /// <param name="ws">The worksheet of the range, or null when the worksheet does not exist.</param>
        /// <returns>The range, or null when the address has no sheet or the worksheet does not exist.</returns>
        IXLRange? RangeFromFullAddress(String rangeAddress, out IXLWorksheet? ws);

        IXLRanges Ranges(String ranges);

        /// <summary>
        /// Force recalculation of all cell formulas.
        /// </summary>
        void RecalculateAllFormulas();

        /// <summary>
        /// Force recalculation of all cell formulas. A cell that fails doesn't stop the recalculation.
        /// </summary>
        /// <remarks>
        /// <para>
        /// A cell fails when its formula is part of a cycle, can't be parsed, exceeds the calculation memory
        /// limit or throws an exception during evaluation (e.g. an unsupported feature). A failed cell keeps its
        /// previous value and is added to <paramref name="failed"/>. The failed cell is marked as calculated, so
        /// reading its value returns the kept value until the cell is marked dirty again.
        /// </para>
        /// <para>
        /// A formula that depends on a failed cell is calculated with the kept value of the failed cell. For
        /// a cycle, only the cell where the cycle is detected fails, other cells of the cycle use its kept value.
        /// </para>
        /// <para>A cancellation of the calculation still throws.</para>
        /// </remarks>
        /// <param name="failed">A collection the cells that failed are added to.</param>
        /// <exception cref="ArgumentNullException"><paramref name="failed"/> is <c>null</c>.</exception>
        void RecalculateAllFormulas(ICollection<IXLCell> failed);

        /// <summary>
        /// Force recalculation of all cell formulas, as <see cref="RecalculateAllFormulas(ICollection{IXLCell})"/>
        /// does. Also add each formula cell that uses an operation of a <see cref="XLCalculationRecordKind"/> to
        /// <paramref name="records"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// A cell gets at most one record of each kind. A failed cell keeps the records of the operations before
        /// the failure. The records of an array formula go to the first cell of its range. The records don't change
        /// the values of the cells. An exception from a collection stops the recalculation and goes to the caller.
        /// A collection must not read the value of another formula cell during the recalculation: the read
        /// calculates that cell outside the recalculation, so the cell gets no record.
        /// </para>
        /// <para>
        /// The records cover the operations of the calculation engine, e.g. a conversion of a text with a digit to
        /// a number. The other overloads don't collect records.
        /// </para>
        /// </remarks>
        /// <param name="failed">A collection the cells that failed are added to.</param>
        /// <param name="records">A collection the records are added to.</param>
        /// <exception cref="ArgumentNullException"><paramref name="failed"/> or <paramref name="records"/> is <c>null</c>.</exception>
        void RecalculateAllFormulas(ICollection<IXLCell> failed, ICollection<XLCalculationRecord> records);

        /// <summary>
        /// Recalculate only the dirty cell formulas.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Unlike <see cref="RecalculateAllFormulas()"/>, the method doesn't mark all formulas as dirty first.
        /// A set of a value or a formula marks the formulas that read the cell as dirty, so the method
        /// calculates only them. A new formula is dirty.
        /// </para>
        /// <para>
        /// A change that doesn't set a value doesn't mark a formula as dirty, e.g. a hide of a row for
        /// <c>SUBTOTAL</c>, or a new time for <c>NOW</c>. Use <see cref="MarkDependentsDirty"/> and
        /// <see cref="IXLCell.InvalidateFormula"/> for these.
        /// </para>
        /// <para>
        /// A set of a value marks the formulas that read the cell only when the dependency tree exists. Each
        /// calculation builds the tree, this method included, and so does <see cref="MarkDependentsDirty"/>.
        /// A loaded workbook has no tree. Call this method once after a load and before the first change.
        /// </para>
        /// </remarks>
        void RecalculateDirtyFormulas();

        /// <summary>
        /// Recalculate only the dirty cell formulas. A cell that fails doesn't stop the recalculation.
        /// </summary>
        /// <remarks>
        /// <para>
        /// The method calculates the same formulas as <see cref="RecalculateDirtyFormulas()"/>. A failed cell
        /// is handled as in <see cref="RecalculateAllFormulas(ICollection{IXLCell})"/>.
        /// </para>
        /// <para>
        /// A failed cell is marked as calculated. A later recalculation of dirty formulas doesn't report the
        /// cell again until the cell becomes dirty, e.g. for a cycle that still exists. Each
        /// <see cref="RecalculateAllFormulas(ICollection{IXLCell})"/> reports it again.
        /// </para>
        /// </remarks>
        /// <param name="failed">A collection the cells that failed are added to.</param>
        /// <exception cref="ArgumentNullException"><paramref name="failed"/> is <c>null</c>.</exception>
        void RecalculateDirtyFormulas(ICollection<IXLCell> failed);

        /// <summary>
        /// Recalculate only the dirty cell formulas, as <see cref="RecalculateDirtyFormulas(ICollection{IXLCell})"/>
        /// does. Also add each calculated formula cell that uses an operation of a <see cref="XLCalculationRecordKind"/>
        /// to <paramref name="records"/>.
        /// </summary>
        /// <remarks>
        /// The records follow the rules of
        /// <see cref="RecalculateAllFormulas(ICollection{IXLCell}, ICollection{XLCalculationRecord})"/>. A formula
        /// that isn't dirty isn't calculated, so it gets no record.
        /// </remarks>
        /// <param name="failed">A collection the cells that failed are added to.</param>
        /// <param name="records">A collection the records are added to.</param>
        /// <exception cref="ArgumentNullException"><paramref name="failed"/> or <paramref name="records"/> is <c>null</c>.</exception>
        void RecalculateDirtyFormulas(ICollection<IXLCell> failed, ICollection<XLCalculationRecord> records);

        /// <summary>
        /// Mark as dirty each cell formula that reads a cell of the <paramref name="range"/>, directly or through
        /// other formulas. The next recalculation of dirty formulas calculates them.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Use the method after a change that doesn't set a value of a cell, e.g. after a hide of a row for
        /// <c>SUBTOTAL</c>. A formula in the <paramref name="range"/> stays as it is. If the dependency tree
        /// doesn't exist (e.g. in a loaded workbook before the first calculation), the method builds it.
        /// </para>
        /// <para>
        /// The mark doesn't go past a formula that is already dirty. After <see cref="IXLCell.InvalidateFormula"/>,
        /// call the method for the cell (<c>cell.AsRange()</c>), so that the formulas that read the cell become
        /// dirty too.
        /// </para>
        /// </remarks>
        /// <param name="range">The range whose dependent formulas become dirty.</param>
        /// <exception cref="ArgumentNullException"><paramref name="range"/> is <c>null</c>.</exception>
        /// <exception cref="ArgumentException"><paramref name="range"/> is not a valid range of a worksheet of the workbook.</exception>
        void MarkDependentsDirty(IXLRangeBase range);

        /// <summary>
        ///   Saves the current workbook.
        /// </summary>
        void Save();

        /// <summary>
        ///   Saves the current workbook and optionally performs validation
        /// </summary>
        void Save(Boolean validate, Boolean evaluateFormulae = false);

        void Save(SaveOptions options);

        /// <summary>
        ///   Saves the current workbook to a file.
        /// </summary>
        void SaveAs(String file);

        /// <summary>
        ///   Saves the current workbook to a file and optionally validates it.
        /// </summary>
        void SaveAs(String file, Boolean validate, Boolean evaluateFormulae = false);

        void SaveAs(String file, SaveOptions options);

        /// <summary>
        ///   Saves the current workbook to a stream.
        /// </summary>
        void SaveAs(Stream stream);

        /// <summary>
        ///   Saves the current workbook to a stream and optionally validates it.
        /// </summary>
        void SaveAs(Stream stream, Boolean validate, Boolean evaluateFormulae = false);

        /// <summary>
        /// Saves the current workbook to a stream.
        /// </summary>
        /// <remarks>
        /// <para>
        /// The method makes the package in memory and then copies it to the <paramref name="stream"/>.
        /// The workbook keeps the package in memory as its source, because the next save must start
        /// from the last saved package. The workbook does not read from the
        /// <paramref name="stream"/>. You can dispose the <paramref name="stream"/> after the save.
        /// One case is different: if the workbook was loaded from the <paramref name="stream"/>, the
        /// method changes the package in the <paramref name="stream"/>.
        /// </para>
        /// <para>
        /// If the <paramref name="stream"/> can read, seek and write, a later <see cref="Save()"/>
        /// writes the full workbook to the <paramref name="stream"/> again. A stream that cannot do
        /// one of these operations, for example the response body of a web server, does not change
        /// the target of <see cref="Save()"/>.
        /// </para>
        /// <para>
        /// A save to the same <paramref name="stream"/> again starts at position 0 and removes all
        /// data in the <paramref name="stream"/>. If you dispose the <paramref name="stream"/>, a later
        /// <see cref="Save()"/> throws <see cref="ObjectDisposedException"/>. Then use one of the
        /// <c>SaveAs</c> methods.
        /// </para>
        /// </remarks>
        void SaveAs(Stream stream, SaveOptions options);

        /// <summary>
        /// Searches the cells' contents for a given piece of text
        /// </summary>
        /// <param name="searchText">The search text.</param>
        /// <param name="compareOptions">The compare options.</param>
        /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
        IEnumerable<IXLCell> Search(String searchText, CompareOptions compareOptions = CompareOptions.Ordinal, Boolean searchFormulae = false);

        XLWorkbook SetLockStructure(Boolean value);

        XLWorkbook SetLockWindows(Boolean value);

        XLWorkbook SetUse1904DateSystem();

        XLWorkbook SetUse1904DateSystem(Boolean value);

        /// <summary>
        /// Gets the Excel table of the given name
        /// </summary>
        /// <param name="tableName">Name of the table to return.</param>
        /// <param name="comparisonType">One of the enumeration values that specifies how the strings will be compared.</param>
        /// <returns>The table with given name</returns>
        /// <exception cref="ArgumentOutOfRangeException">If no tables with this name could be found in the workbook.</exception>
        IXLTable Table(String tableName, StringComparison comparisonType = StringComparison.OrdinalIgnoreCase);

        Boolean TryGetWorksheet(String name, [NotNullWhen(true)] out IXLWorksheet? worksheet);

        IXLWorksheet Worksheet(String name);

        IXLWorksheet Worksheet(Int32 position);
    }
}
