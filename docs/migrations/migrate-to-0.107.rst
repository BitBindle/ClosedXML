#############################
Migration from 0.106 to 0.107
#############################

*****************
Target frameworks
*****************

ClosedXML now targets only ``net10.0``. This version does not have the ``netstandard2.0`` and
``netstandard2.1`` builds. .NET Framework, Mono and Unity applications can't use this version.
These applications can stay on version 0.106.

*********************************
Automatic colors of Excel 2010
*********************************

A save now writes ``auto="1"`` for an ``XLColor.Automatic`` color of a data bar or a sparkline.
Before, the save wrote an empty color element, and the file lost the automatic value. A load reads
the element as ``XLColor.Automatic``, as before.

***********************************
Measurement of bold and italic text
***********************************

``DefaultGraphicEngine`` now measures bold and italic text with the bold and italic metrics of the
font. Before, the engine used the regular metrics for all text. ``AdjustToContents`` can now give
different column widths and row heights for bold and italic text. Columns with bold text are
usually wider. If the font does not have the style, the engine uses the nearest style of the font.
For example, it uses the bold style for bold italic text if the font has no bold italic style.

*******************************************************
Copy of a sparkline with source data outside the sheet
*******************************************************

A copy of a cell or a range moves the source data of a sparkline with the cell. When the moved source
data is after the last row or the last column, the copy now does not copy the sparkline. The copy
still copies the value and the format of the cell. Excel shows such a sparkline with a ``#REF!``
reference, and that sparkline shows no data. Before, the copy cleared the target, copied some cells
and then threw ``InvalidOperationException``.

*************
Save a stream
*************

``IXLWorkbook.SaveAs(Stream)`` of a workbook that you loaded from a file or a stream now works with
a stream that cannot read or seek, for example the response body of a web server. Before, it threw
``OpenXmlPackageException``. The method makes the package in memory and then copies it to the
stream. A new workbook already did this.

``SaveAs(Stream)`` now makes the package in memory and then copies it to the stream. The workbook
keeps the package in memory as its source, because the next save must start from the last saved
package. The workbook does not read from the stream. You can dispose the stream after the save.
Before, the stream became the source, and a save after the dispose threw
``ObjectDisposedException``. The copy in memory uses memory that is the size of the saved file.

If the stream can read, seek and write, a later ``Save()`` writes the full workbook to the stream
again. This was also the behavior before. A stream that cannot do one of these operations does not
change the target of ``Save()``. A later ``Save()`` writes to the file or the stream from before.
A save to the same stream again starts at position 0 and removes all data in the stream. If you
dispose the stream, a later ``Save()`` throws ``ObjectDisposedException``. Then use one of the
``SaveAs`` methods.

``SaveAs(String)`` of a workbook that you loaded from a stream now gets the document type from the
file extension. A workbook that you loaded from a file already did this. Before, the save kept the
document type of the stream.

A workbook that you loaded from a stream changes the package in that stream on ``SaveAs`` to the
same stream. This was also the behavior before.

**********************************
Load and save a stream with async
**********************************

``XLWorkbook`` has two new methods:

* ``XLWorkbook.SaveAsAsync(Stream, SaveOptions?, CancellationToken)``. If the stream cannot do one
  of these operations (read, seek or write), the method makes the package in memory. Then it
  copies the package to the stream and flushes the stream with asynchronous calls. Use it for a
  stream that does not allow synchronous writes, for example the response body in ASP.NET Core.
* ``XLWorkbook.LoadAsync(Stream, LoadOptions?, CancellationToken)``. If the stream cannot seek, the
  method first copies the stream to memory with asynchronous calls. Then it reads the copy. A
  workbook from such a copy cannot ``Save()``. Use one of the ``SaveAs`` methods.

If the stream can seek, ``LoadAsync`` reads the stream synchronously. A stream that can seek, but
does not allow synchronous reads, causes an exception. An example is an ASP.NET Core request body
after ``EnableBuffering``. Copy such a stream to a ``MemoryStream`` first.

The work on the package stays synchronous, because the OpenXML SDK has no asynchronous API. The
methods are not on ``IXLWorkbook``, so other implementations of the interface do not break.

**********************
Text from GetValue
**********************

``IXLCell.GetValue<String>()`` and ``IXLCell.TryGetValue<String>()`` now return the text of the
cell without a change. Before, they replaced each escaped character, for example ``_x0041_``, with
the character that it encodes. A load already decodes these escaped characters. The second decode
changed the text. For example, a cell with the text ``a_x0041_`` gave ``aA``. It now gives
``a_x0041_``, the same text as ``IXLCell.GetText()``.

*******************************
Table names in the saved file
*******************************

A save now compares table names without case, as Excel does. A table with the name of an earlier
saved table gets a number in the file. Before, a file could have the names ``Sales`` and ``SALES``,
and Excel had to repair it. Now the second table gets the name ``SALES1``. The table in the workbook
keeps its name.

If the name with the number is a cell reference, the save puts an underscore before the number. For
example, a second table ``AB`` gets the name ``AB_1``, not ``AB1``.

The save does not give a table the name of another table or of a defined name in the workbook. For
example, if the workbook also has a table ``SALES1``, the second table gets the name ``SALES2``.

The save does not change the formulas, the defined names and the pivot sources that use the old
name. Excel compares names without case, so a reference to ``SALES`` goes to the table ``Sales``.
Give each table a name that is unique without case.

*******************************************
Array formulas after an insert or a delete
*******************************************

An insert or a delete of rows, columns or cells now keeps an array formula as an array formula.
The range of the array formula moves with its cells, and the references in the formula change.
Before, each cell of the array formula got a normal formula.

An insert or a delete now also changes the input cells of a data table. If you delete an input
cell, the data table gets ``#REF!`` for that input.

Excel does not allow an insert or a delete that changes only part of an array formula or a data
table. ClosedXML now also stops these operations. Before, ClosedXML did them. The operation throws
``InvalidOperationException`` before it changes the sheet. These operations throw:

* An insert of rows or columns inside the range.
* A delete of some rows or some columns of the range.
* An insert or a delete of cells that moves only some cells of the range.
* An insert that moves the range, or part of it, out of the sheet.

A delete of all rows or all columns of the range does not throw. It deletes the array formula or
the data table.

A delete of a collection of rows or columns checks all parts before the first delete. Some
collections of range rows, range columns or table rows have parts that overlap. Other collections
have range rows with different columns, or range columns with different rows. For these
collections, the delete throws if an array formula or a data table is in the cells that the delete
can move.

``IXLRange.Transpose`` and ``IXLTable.Resize`` do more than one insert or delete. They now check all
steps before the first change to the sheet. If a step changes only part of an array formula or a
data table, the method throws. Then the sheet does not change.

``IXLRange.Transpose`` also throws when it moves cells of an array formula or a data table. The
transpose cannot change the range of the formula. A cell on the diagonal does not move. With
``XLTransposeOptions.ReplaceCells``, a transpose clears the cells next to the range. It throws if
the clear removes only part of an array formula or a data table.

*****************************************
Clear of part of an array formula
*****************************************

``IXLRangeBase.Clear`` and ``IXLCell.Clear`` now throw ``InvalidOperationException`` when they
remove the contents of only part of an array formula or a data table. The method throws before it
changes the sheet. Before, the clear removed the formula from the cleared cells only. The other
cells kept the array formula with the old range.

These clears do not throw:

* A clear of all cells of the array formula or the data table. It removes the formula.
* A clear without ``XLClearOptions.Contents``, for example a clear of the formats.

A clear of a collection, for example ``IXLCells.Clear`` or ``IXLRanges.Clear``, checks all parts
before the first change. The parts together must contain all cells of the array formula or the
data table. If they do not, the clear throws, and the sheet does not change.

Some other methods clear cells, so they can also throw. ``IXLCell.CopyFrom(IXLRangeBase)``,
``IXLColumn.CopyTo(IXLColumn)`` and ``IXLRow.CopyTo(IXLRow)`` throw before they change the sheet.
``IXLTable.ShowHeaderRow`` with the value false can also throw. It throws after it writes the
names of blank header cells.

*************************************
One-cell array formula in cell A1
*************************************

A one-cell array formula in cell A1 now has a range. ``IXLCell.FormulaReference`` gives
``A1:A1``. The save of the workbook writes the formula. Before, ``FormulaReference`` gave null.
The save threw ``InvalidOperationException``.

``IXLCell.FormulaReference`` does not accept null for an array formula or a data table. If you set
it to null, it throws ``ArgumentException``. Before, it changed the range of the formula to
``A1:A1``. To remove an array formula, set the value of its range.

************************
Copy of an array formula
************************

A copy of cells now keeps an array formula. The copy has one array formula with the same size at
the new position. The relative references of the formula move with the range. This applies to
``IXLWorksheet.CopyTo``, to ``CopyTo`` of a range, a row or a column, and to ``IXLCell.CopyFrom``
and ``IXLCell.CopyTo``. Before, each copied cell got a normal formula with the text of the array formula.

A copy of only part of an array formula now throws ``InvalidOperationException``. The copy throws
before it changes the target. Excel does not allow this copy.

A copy of one cell to a cell of an array formula or a data table with more than one cell now throws
``InvalidOperationException``. A copy of only the cell format does not throw.

A copy that puts an array formula with more than one cell over a table now throws
``InvalidOperationException``. A copy that puts an array formula in a merged cell also throws.
Excel does not allow these array formulas. The copy throws before it changes the target.

A copy of a sheet and ``IXLTable.CopyTo`` copy a table with its cells. A table from a file can have an
array formula with more than one cell, because the load does not check the table. These copies keep
that array formula in the new table.

A copy of a sheet now keeps a data table. A copy of a range or a cell copies only the values of a
data table, as Excel does. Before, these copies threw ``ParsingException``.

A copy of a sheet now merges the cells with ``IXLRangeBase.Merge(false)``. The copy keeps the
contents and the format of the merged cells. Before, the merge cleared the cells after the first
cell and changed their borders.

*****************************************
Table over an array formula
*****************************************

These methods now throw ``InvalidOperationException`` when the table range has cells of an array
formula or a data table with more than one cell:

* ``IXLRange.CreateTable`` and ``IXLRangeColumn.CreateTable``.
* ``IXLTable.Resize``.
* ``IXLTable.CopyTo``.
* ``IXLCell.InsertTable`` with ``createTable`` true. It throws before it writes the data.

Excel does not allow these formulas in a table. Before, ClosedXML created the table. An array
formula in one cell does not cause the exception.

The load of a workbook does not do this check. A file with such a table still loads.

``IXLRangeBase.FormulaArrayA1`` now accepts an array formula in one cell of a table. Before, it
threw ``InvalidOperationException``. It still throws for an array formula with more than one cell.

*****************************************
Merge over an array formula
*****************************************

``IXLRangeBase.Merge`` now throws ``InvalidOperationException`` when the range has cells of an
array formula or a data table. This also applies to an array formula in one cell. The method throws
before it changes the sheet. Excel does not allow these formulas in merged cells. Before, the merge
cleared the formula from some cells only. The other cells kept the formula with the old range.

``IXLRangeBase.Merge(false)`` does not do this check. It does not clear cells. The load of a
workbook and the copy of cells use it, so a file with such a merged range still loads.

*****************
Equality methods
*****************

``XLColor.Equals(object)`` now returns false for null and for an object that is not an ``XLColor``.
Before, it threw an exception. The ``IEquatable<T>.Equals`` methods of a cell fill, of a differential
number format and of a differential protection now return false for null. Before, they threw
``NullReferenceException``.

*****************************************************
Copy of a range after the last row or the last column
*****************************************************

A range copy now throws ``InvalidOperationException`` when the copy of a used cell or a merged
range is after the last row or the last column of the sheet. The copy throws before it changes the
target. Excel also does not paste a copy that does not fit. Before, the copy cleared the target,
copied some cells and then threw ``ArgumentOutOfRangeException``.

A copy that has only empty cells after the last row or the last column still copies the used cells.

A range copy now copies only the part of a data validation that is in the source. Before, the copy
also put the data validation on cells outside the target. When the copied data validation was after
the last row or the last column, the copy threw ``ArgumentOutOfRangeException``.

*********
New lines
*********

``AddNewLine`` of a comment, of rich text and of a header or footer item now adds a line feed
(``\n``). ``IXLComment.AddSignature`` also adds a line feed after the author. Before, these
methods added ``Environment.NewLine``, which is ``\r\n`` on Windows. Excel uses a line feed for a
new line.

A text value with a line feed now gets the wrap text format on all operating systems. Before, a
value got the format only if it contained ``Environment.NewLine``. On Windows, a value with only
``\n`` did not get it.

A save now writes a line feed in the name of a table column as ``_x000A_`` and a carriage return as
``_x000D_``. A load decodes all escape sequences in the name. Before, the save and the load changed
``Environment.NewLine``, and the name was different on Windows and Linux. The load also changed the
text ``_x000a_`` in a name to a new line.

With a header row, a table field gets the text of its header cell during the load. If the decoded
names of a table are not unique, the load keeps the names from the file. An earlier version wrote a
name such as ``_x0041_`` without an escape, so the load can still read such a file.

``IXLTable.Field(String)`` and the other methods that find a table field by its name no longer
change ``\r\n`` to ``\n`` in the name. Use the same text as in the header cell.

**********************
Text in the saved file
**********************

A save now writes a carriage return as ``_x000D_``. Before, it wrote the character, and a load
changed it to a line feed. Excel also writes ``_x000D_``.

A save now encodes the text of inline strings, rich text, phonetic runs, comments and the cached
text of formulas. Before, the save encoded only the shared strings. Text such as ``_x0041_`` now
stays the same after a save and a load. A control character, for example U+0001, no longer causes
an exception.

A load now decodes the escaped characters of a comment. Before, a comment showed the escape text,
for example ``_x000D_``.

A save now also encodes the formula text of a cell and the author of a comment. A load decodes
them, because the file format defines both as ``ST_Xstring``. Before, a carriage return in a
formula became a line feed, and a control character caused an exception. An escape sequence in a
loaded formula, for example ``_x000D_``, now gives the character. Before, the formula kept the
escape text.

A save and a load now do the same for the formulas of defined names, data validations and
conditional formats. A save also encodes the calculated column formula and the totals row formula
of a table column. A load decodes the totals row formula. The file format defines these formulas as
``ST_Formula``, which is an ``ST_Xstring``. Before, a carriage return in these formulas became a
line feed, and a formula kept the escape text from a file. A control character in a data
validation formula no longer causes an exception on save.

A save and a load now also encode and decode these text attributes, because the file format
defines them as ``ST_Xstring``:

* The input title, the input message, the error title and the error message of a data validation.
* The text of a conditional format rule, for example a rule that finds cells that contain a text.

Before, an escape sequence from a file, for example ``_x000D_``, stayed as the escape text. A control
character in these texts caused an exception on save. Now the save encodes it.

A load no longer changes the line ends of an inline string or of a comment. The text stays the same
as in the file, as for a shared string. Excel uses a line feed (``\n``) for a new line in a cell.
Before, the load changed each line end to ``Environment.NewLine``. On Windows, a line feed became
``\r\n``. On Linux, ``\r\n`` became a line feed. To get ``Environment.NewLine``, replace the line
feeds in the text after the load.

********************
Nullable annotations
********************

More files of ClosedXML now have nullable reference types. The public API now shows where a value
can be null, and your compiler can show new warnings. Most annotations do not change the behavior at
runtime. Some members now throw a specified exception where before they threw
``NullReferenceException`` or returned null. The list below gives these members:

* ``IXLAutoFilter.Range`` is now ``IXLRange?``. It is null when the worksheet has no autofilter.
  For an autofilter without a range, these members now throw ``InvalidOperationException``.
  Before, they threw ``NullReferenceException``:

  * ``HiddenRows`` and ``VisibleRows``.
  * ``Sort`` and ``Reapply`` of an enabled autofilter.
  * Each ``IXLFilterColumn`` filter method with ``reapply`` set to ``true`` (the default), because
    it calls ``Reapply``.
  * ``AboveAverage`` and ``BelowAverage``, also with ``reapply`` set to ``false``.

  ``Clear`` of an enabled autofilter without a range does not throw now.
* ``IXLConditionalFormat.Values`` is now ``XLDictionary<XLFormula?>``. A color scale with the lowest
  value or the highest value has a null value at that position.
* ``IXLAddress.Worksheet`` is now ``IXLWorksheet?``. ``ToString(XLReferenceStyle, Boolean)``,
  ``ToStringRelative(Boolean)`` and ``ToStringFixed(XLReferenceStyle, Boolean)`` with
  ``includeSheet`` set to ``true`` on an address without a worksheet now throw
  ``InvalidOperationException``. Before, they threw ``NullReferenceException``.
* ``IXLRanges.GetIntersectedRanges(IXLAddress)`` with an address without a worksheet now throws
  ``ArgumentException``. Before, it threw ``ArgumentNullException``.
* ``IXLHFItem.AddText``, ``IXLHFItem.AddNewLine`` and ``IXLPrintAreas.Add`` on the page setup of
  ``XLWorkbook.PageOptions`` now throw ``InvalidOperationException``, because that page setup has
  no worksheet. Before, they threw ``NullReferenceException``.
* ``IXLWorkbook.Cell(String)`` is now ``IXLCell?`` and ``IXLWorkbook.Range(String)`` is now
  ``IXLRange?``. They return null when the name or the address does not refer to a cell or a range.
* ``IXLWorkbook.RangeFromFullAddress`` now returns ``IXLRange?`` and has the parameter
  ``out IXLWorksheet? ws``. Both are null when the address has no sheet or the worksheet does not
  exist.
* ``XLWorkbook.CellFromFullAddress`` now returns ``IXLCell?`` and has the parameter
  ``out IXLWorksheet? ws``. Both are null when the address has no sheet or the worksheet does not
  exist.
* ``IXLWorkbook.TryGetWorksheet`` now has the parameter
  ``[NotNullWhen(true)] out IXLWorksheet? worksheet``.
* The setter of ``IXLWorkbook.Protection`` now throws ``ArgumentNullException`` for null. Before,
  it threw ``NullReferenceException``.
* ``IXLPicture.TopLeftCell`` and ``IXLPicture.BottomRightCell`` are now ``IXLCell?``. A picture
  that has no anchor at a position returns null. Before, they threw ``NullReferenceException``.
  ``IXLPicture.GetOffset`` for a position without an anchor now throws
  ``InvalidOperationException``.
* ``IXLPictures.TryGetPicture`` now has the parameter ``[NotNullWhen(true)] out IXLPicture? picture``.
* ``IXLCustomProperties.Add<T>`` with a null value and ``IXLCustomProperty.Value = null`` now throw
  ``ArgumentNullException``.
* The ``XLFormula(String?)`` constructor and the setter of ``XLFormula.Value`` accept null. Null
  gives an empty value. This was already the behavior at runtime.
* ``new XLFormula().Value`` is now an empty string. Before, it was null.
* ``IXLPivotTable.Title``, ``Description``, ``ColumnHeaderCaption``, ``RowHeaderCaption`` and
  ``ErrorValueReplacement`` are now ``String?``. Their ``Set`` methods accept ``String?``.
* ``IXLPivotValue.BaseFieldName`` is now ``String?``. ``SetBaseFieldName`` accepts ``String?``.
* ``IXLPivotField.CustomName`` is now ``String?``. It is null when the field has no custom name.
  ``SetCustomName`` accepts ``String?``. Null for the data field (``Σ Values``) throws
  ``ArgumentNullException``.
* ``IXLPivotFields.Add(String, String?)`` accepts a null custom name. Null adds the field without a
  custom name.
* ``IXLSparklineGroup.DateRange`` is now ``IXLRange?`` and ``SetDateRange`` accepts ``IXLRange?``.
  ``IXLSparklineGroup.GetSparkline`` and ``IXLSparklineGroups.GetSparkline`` now return
  ``IXLSparkline?``.
* ``IXLSparklineGroups.Add(String, String)``, ``IXLSparklineGroup.Add(String, String)`` and
  ``IXLPrintAreas.Add(String)`` with an address or a name that does not refer to a range now throw
  ``ArgumentException``.
* ``IXLTheme.ResolveThemeColor`` with a value that is not a member of ``XLThemeColor`` now throws
  ``ArgumentOutOfRangeException``.
* The members that find a used part of a range now return a nullable type. They return null when
  the range has no used cells:

  * ``FirstCellUsed`` and ``LastCellUsed`` of ``IXLRangeBase`` (all overloads) return ``IXLCell?``.
  * ``IXLRange.RangeUsed()`` returns ``IXLRange?``.
  * ``FirstColumnUsed`` and ``LastColumnUsed`` of ``IXLRange`` return ``IXLRangeColumn?``.
  * ``FirstRowUsed`` and ``LastRowUsed`` of ``IXLRange`` return ``IXLRangeRow?``, and of
    ``IXLTableRange`` return ``IXLTableRow?``.
  * ``ColumnUsed`` of ``IXLColumn`` and ``IXLRangeColumn`` returns ``IXLRangeColumn?``. ``RowUsed``
    of ``IXLRow`` and ``IXLRangeRow`` returns ``IXLRangeRow?``.

* ``FirstColumn``, ``LastColumn``, ``FirstRow`` and ``LastRow`` of ``IXLRange`` and ``FirstRow``
  and ``LastRow`` of ``IXLTableRange`` had an optional predicate. Each method is now two overloads.
  The overload without a parameter returns a non-null value. The overload with a predicate returns
  a nullable value, which is null when no column or row matches the predicate. Code that calls these
  methods compiles without changes, but you must compile it again, because the binary signature
  changed.
* ``IXLRange.FindColumn`` returns ``IXLRangeColumn?`` and ``IXLRange.FindRow`` returns
  ``IXLRangeRow?``. They return null when no column or row matches the predicate.
* ``IXLRangeBase.GetDataValidation()`` returns ``IXLDataValidation?``. It returns null when no data
  validation rule has exactly the address of the range. The documentation said before that the
  method creates a new rule. It does not.
* ``IXLRangeBase.Shrink()`` and ``IXLRangeBase.Shrink(Int32)`` return ``IXLRangeBase?``. They
  return null when no cells remain.
* ``IXLRangeBase.Intersection`` returns ``IXLRangeAddress?``, and accepts a null range.
  ``IXLRangeBase.Union`` and ``IXLRangeBase.Difference`` accept a null range. Null gives the cells
  of the range that match the predicate of the range.
* ``IXLRangeBase.AddToNamed(String, XLScope, String?)`` accepts a null comment.
* The optional predicate parameters of these members are now ``Func<...>?``: ``Columns``,
  ``ColumnsUsed``, ``Rows``, ``RowsUsed``, ``FirstColumnUsed``, ``LastColumnUsed``,
  ``FirstRowUsed`` and ``LastRowUsed`` of ``IXLRange``, ``Rows``, ``RowsUsed``, ``FirstRowUsed``
  and ``LastRowUsed`` of ``IXLTableRange``, and ``SurroundingCells``, ``Intersection``, ``Union``
  and ``Difference`` of ``IXLRangeBase``. Null gives all the items. This was already the behavior at
  runtime.
* ``IXLRangeBase.Intersects(String)`` on a worksheet with a name that is not an address or a
  defined name now throws ``ArgumentException``. Before, it threw ``NullReferenceException``.
  ``IXLRange.Cell(String)`` with such a name now throws ``ArgumentException``. Before, it returned
  null.
* ``IXLWorksheet.Sort()`` on a worksheet without used cells now throws
  ``InvalidOperationException``. Before, it threw ``NullReferenceException``.
* ``IXLCell.FormulaA1`` and ``IXLCell.FormulaR1C1`` accept null. Null removes the formula. This was
  already the behavior at runtime.
* ``IXLCell.FormulaReference`` is now ``IXLRangeAddress?``, ``IXLCell.Sparkline`` is now
  ``IXLSparkline?`` and ``IXLCell.MergedRange()`` returns ``IXLRange?``. They are null when the
  cell has no array formula, no sparkline or no merged range.
* ``IXLCell.TryGetValue<T>`` now has the parameter ``[MaybeNullWhen(false)] out T value``.
* ``IXLCell.GetFormattedString(CultureInfo?)`` accepts a null culture. Null gives the current
  culture. This was already the behavior at runtime.
* ``IXLCell.InsertData`` (all overloads) returns ``IXLRange?`` and accepts null data. It returns
  null when the data is null or a string.
* ``IXLCell.InsertTable`` with a ``DataTable`` returns ``IXLTable?`` and accepts a null data table.
  It returns null when the data table is null or has no columns. All ``InsertTable`` overloads
  accept a null table name. Null gives a generated name.
* ``IXLTable.DataRange`` is now ``IXLTableRange?``. It is null when the table has no data rows.
  ``IXLTable.HeadersRow()`` and ``IXLTable.TotalsRow()`` return ``IXLRangeRow?``. They are null when
  the table does not show the row.
* ``IXLTable.AppendData`` (all overloads) returns ``IXLRange?``. It returns null when the data is
  empty or a string. The overloads with ``IEnumerable`` and ``IEnumerable<T>`` accept null data, and
  return null for it. The overload with a ``DataTable`` does not accept null. Null now throws
  ``ArgumentNullException``. Before, it threw ``NullReferenceException``. This is also true for
  ``IXLTable.ReplaceData(DataTable, Boolean)``.
* ``AppendData``, ``ReplaceData`` and ``Sort`` of a table without data rows now throw
  ``InvalidOperationException``. Before, they threw ``NullReferenceException``.
* ``IXLTableField.HeaderCell`` and ``IXLTableField.TotalsCell`` are now ``IXLCell?``. They are null
  when the table does not show the row. ``IXLTableField.TotalsRowLabel`` is now ``String?``. The
  setter of ``TotalsRowLabel``, and the getters and the setters of ``TotalsRowFormulaA1`` and
  ``TotalsRowFormulaR1C1``, now throw ``InvalidOperationException`` when the table does not show
  the totals row. Before, they threw ``NullReferenceException``.
* ``IXLTables.TryGetTable`` now has the parameter ``[NotNullWhen(true)] out IXLTable? table``.
* ``XLHyperlink.ExternalAddress`` is now ``Uri?`` and ``XLHyperlink.InternalAddress`` is now
  ``String?``. Each is null when the hyperlink does not have that kind of address. Their setters now
  throw ``ArgumentNullException`` for null. The setter of ``XLHyperlink.Tooltip`` accepts null. Null
  gives an empty tooltip. The tooltip of a hyperlink from ``IXLCell.CreateHyperlink()``, and from
  ``IXLCell.GetHyperlink()`` on a cell without a hyperlink, is now an empty string. Before, it was
  null.
* ``XLHelper.IsValidA1Address`` and ``XLHelper.IsValidRCAddress`` accept null. Null gives false.
* The implicit conversion from ``String?`` to ``XLCellValue`` accepts null. Null gives a blank
  value. This was already the behavior at runtime. ``XLCellValue.FromObject`` accepts null.
  ``XLCellValue.TryGetText`` now has the parameter ``[NotNullWhen(true)] out String? value``, and
  ``XLCellValue.TryConvert`` for ``Blank`` has the parameter ``[NotNullWhen(true)] out Blank? value``.
  ``XLCellValue.Equals(Blank?)`` and ``XLCellValue.Equals(String?)`` accept null.
  ``Equals(Blank?)`` returns true for a blank value, also for null. ``Equals(String?)`` returns
  false for null.

These bugs are fixed:

* ``HighestValue`` of a color scale removed the value of ``Minimum``.
* A copy of a color scale with the lowest value or the highest value threw
  ``NullReferenceException``.
* ``IXLPivotFields.Add(String, String?)`` with a null custom name threw
  ``InvalidOperationException``. A copy of a pivot table with a field that has no custom name threw
  the same exception, because the copy uses this method.
* ``IXLPivotField.SetCustomName(null)`` threw ``ArgumentException`` when another field had no
  custom name.
* A new table on a worksheet with an enabled autofilter that has no range threw
  ``NullReferenceException``. A copy of that worksheet threw the same exception.
* ``IXLRanges.AddToNamed(String, XLScope)`` ignored the scope and always added a workbook name.
* ``IXLTable.Resize(IXLRangeAddress)`` ignored the address and used the current address of the
  table.
* A resize of a table with a totals row to more rows gave the table one row fewer. The totals row
  was not in the last row of the new range.
* ``FirstRowUsed`` and ``LastRowUsed`` of ``IXLTableRange`` without a predicate returned a row that
  was not valid when the table range had no used rows. They now return null.
* A new name, or a new totals row function, for a table with a hidden totals row threw
  ``NullReferenceException`` when a field had a totals row function.
* ``Cells(false)`` of a column, of a row and of a worksheet without used cells threw
  ``NullReferenceException``. They now return no cells.
* ``Sort()`` and ``SortLeftToRight()`` of an entire-row or entire-column range without used cells
  threw ``NullReferenceException``. When the worksheet also had no used cells, they threw
  ``ArgumentOutOfRangeException``. They now do nothing.
* A copy of a cell with a sparkline that has no source data threw ``NullReferenceException``. The
  copy now does not include the sparkline.
* ``IXLTable.AsDynamicEnumerable()`` and ``IXLTable.AsNativeDataTable()`` of a table without data
  rows threw ``NullReferenceException``. They now return no rows.
* A load of a file without a required part, element or attribute threw
  ``NullReferenceException``. It now throws ``PartStructureException``. A load of a file with a
  reference that is not valid, for example the ``ref`` of an autofilter, a merged range, a table, a
  comment or a pivot table location, the ``sqref`` of a conditional format, or the ``activeCell``
  of a selection, also throws ``PartStructureException``.
* A load of a file with an autofilter ``dynamicFilter`` without a ``val`` threw
  ``NullReferenceException``. The ``DynamicValue`` of the filter column is now ``Double.NaN``. A
  save does not write ``val`` for a ``DynamicValue`` that is ``Double.NaN``.
* A save of a file with a count ``vt:variant`` in ``docProps/app.xml`` without a ``vt:i4`` child
  threw ``NullReferenceException``. The save now adds the ``vt:i4`` element.
* A load of a file with a row page break or a column page break without an ``id`` threw
  ``NullReferenceException``. The load now skips the page break.
* A load of a file with an autofilter ``filter`` or ``customFilter`` without a ``val`` threw
  ``NullReferenceException``. The filter now compares with an empty text.
* A load of a picture with a two-cell anchor, ``editAs="absolute"`` and shape properties without
  an offset threw ``NullReferenceException``. The load now keeps the picture and does not move it.
* A save of a worksheet with comments and a VML drawing part that is not valid XML threw
  ``NullReferenceException``. The save now writes the comments without the old VML content.
* A save after the delete of a worksheet with a pivot table threw
  ``ArgumentOutOfRangeException``.
