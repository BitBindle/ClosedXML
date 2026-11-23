#############################
Migration from 0.106 to 0.107
#############################

*****************
Target frameworks
*****************

ClosedXML now targets only ``net10.0``. This version does not have the ``netstandard2.0`` and
``netstandard2.1`` builds. .NET Framework, Mono and Unity applications can't use this version.
These applications can stay on version 0.106.

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

*****************
Equality methods
*****************

``XLColor.Equals(object)`` now returns false for null and for an object that is not an ``XLColor``.
Before, it threw an exception. The ``IEquatable<T>.Equals`` methods of a cell fill, of a differential
number format and of a differential protection now return false for null. Before, they threw
``NullReferenceException``.

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
