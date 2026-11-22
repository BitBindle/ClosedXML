#############################
Migration from 0.105 to 0.106
#############################

**************
Number formats
**************

Format is always there
----------------------

Number format ```IXLNumberFormatBase.Format``` previously returned an empty
string for predefined formats. It now  returns actually used predefined format
code instead.

Number format id
----------------

The number format setters (```IXLNumberFormatBase.NumberFormatId```,
```IXLNumberFormat.SetNumberFormatId(int)``` and ```IXLPivotValueFormat.SetNumberFormatId(int)```)
now throw an ```ArgumentOutOfRangeException``` when supplied number format id is
not a predefined format id from ```XLPredefinedFormat```.

************
IXLAlignment
************

The ``IXLAlignment.TextRotation`` now throws ``ArgumentOutOfRangeException`` on invalid rotation
instead of original ``ArgumentException``.

The ``IXLAlignment.Indent`` now throws ``InvalidOperationException`` instead of original
``ArgumentException`` when trying to set indent on format with horizontal alignment that doesn't
support indent.

*************
IXLWorksheets
*************

The method ``IXLWorksheets.Worksheet(string sheetName)`` now throws ``KeyNotFoundException`` when
sheet is not found, instead of original ``ArgumentException``.

Generally speaking, get-by-name methods of all unique-name collections (e.g., worksheets, styles,
table fields, etc.) should throw ``KeyNotFoundException`` when an item is not found. This change
aligns the behavior of the ``IXLWorksheets`` with the rest of the API.

*************
Defined names
*************

A property setter ``IXLDefinedName.RefersTo`` now throws an ``ArgumentException``
when trying to set an empty/whitespace-only value.

*************
Sparklines
*************

Property ``IXLSparkline.IsValid`` has been removed. The ``IXLSparkline`` is now an API object and
the sparkline data should be always valid. Reminder: API objects created before a structural
change are no longer valid after structural change.

Property ``IXLSparkline.Location`` setter now throws ``ArgumentException`` on a cell from different
sheet instead of original ``InvalidOperationException``.

************
IXLRangeBase
************

An obsolete method ``IXLRangeBase.SetDataValidation`` has been removed. Use ``GetDataValidation()``
to access the existing rule, or ``CreateDataValidation()`` to create a new one.

*******
IXLCell
*******

An obsolete method ``IXLRangeBase.SetDataValidation`` has been removed. Use ``GetDataValidation()``
to access the existing rule, or ``CreateDataValidation()`` to create a new one.

*******
XLColor
*******

There is now a ``XLColor.Automatic`` color. It represents the automatic color from the color
picker. The actual value of a color is determined by an application based on a use, generally
either a black (e.g. font name) or white (e.g. fill). There is also a new enum member
``XLColorType.Automatic`` for the color.

*****************
XLWorkbook.Style
*****************

The ``XLWorksheet.Style`` now represents the default format of a workbook. That means all cells
without explicit format will use the default format. Previously, it was only storing style that
was used for newly created worksheets.

The ``XLWorksheet.DefaultStyle`` static property has been removed. The style is now an API object
that only access data structures behind. Use ``XLWorksheet.Style`` to set format of elements back
to the default format of a workbook (e.g. ``ws.Cell("A1").Style = wb.Style``).

Default color of fill and border colors ``IXLStyle`` elements is now ``XLColor.Automatic``.
Previously, it was (``XLColor.FromIndex(64)``).

****************************
Clearing conditional formats
****************************

Partial clearing of conditional format (e.g. clearing `B2` of conditional format for area ``A1:C3``)
now correctly splits the format of into individual remaining areas (e.g. into ``A1:A3 B1 C1:C3 B3``).

Originally, the partial clearing only worked for some situations. Concretely, it only worked when
the area was fully horizontally/vertically split, but not for corners (``A1:B2`` without a corner) or
center (``A1:C3`` without ``B2``).

Reminder: the clearing of a conditional format is done through the Clear method with appropritate
option, e.g. ``ws.Range("A2:C5").Clear(XLClearOptions.ConditionalFormats)``.

***************************
IXLConditionalFormat.Ranges
***************************

The ``IXLConditionalFormat.Ranges`` property had its type changed from ``IXLRanges`` to ``IEnumerable<IXLRange>``.
If you need to set ranges, use a newly added setter, e.g. ``conditionalFormat.Ranges = ws.Ranges("A1:C5,G1:J5")``.

All conditional formats now must apply to some range. They can't be empty (=with no range).

************************
Conditional text formats
************************

The methods ``IXLConditionalFormat.WhenContains``, ``WhenNotContains``, ``WhenStartsWith`` and
``WhenEndsWith`` keep the text as a literal value:

* The methods do not trim the text.
* A leading ``=`` is a part of the text. Before, ``WhenContains("=A1")`` removed the ``=`` and
  found the text ``A1``. Now it finds the text ``=A1``.

The value that ClosedXML loads from a file is also a literal text.

***********
Table names
***********

The setter ``IXLTable.Name`` now throws an ``ArgumentException`` for a name that Excel does not
accept:

* The name contains whitespace, for example ``My table``.
* The name is a cell reference, for example ``A1``, ``May2019`` or ``R1C2``.
* The name is the name of a defined name in the workbook or in one of its worksheets. The compare
  ignores case.

Before, ClosedXML kept such a name in memory, but the save removed the whitespace. Structured
references and pivot sources that used the name then pointed to a table that did not exist. A file
with a cell reference or a defined name as a table name needed a repair in Excel.

The ``CreateTable(name)``, ``AsTable(name)`` and ``InsertTable(data, name)`` methods use the same
rules. For a name that is a cell reference, the ``InsertTable`` overload for a ``DataTable`` throws
an ``InvalidOperationException``, as before.

The load and a copy of a table keep a table name from a file as it is. The save still removes the
whitespace from such a name.

*******************************************
Where, Skip and Take on rows and columns
*******************************************

The ``IXLColumns``, ``IXLRows``, ``IXLRangeColumns`` and ``IXLRangeRows`` interfaces have new
``Where``, ``Skip`` and ``Take`` methods. Each method gives a collection of the same type, for
example ``ws.ColumnsUsed().Skip(1).AdjustToContents()``.

* A call such as ``columns.Where(c => ...)`` now uses the new method and not LINQ. The result is
  a new collection. The methods do not defer the work.
* The result has the order of the enumeration: the column or row number.
* The result of ``ws.Columns()`` or ``ws.Rows()`` with a filter does not represent all columns or
  rows of the sheet. The ``Width``, ``Height`` and ``Delete`` members change only the columns or
  rows in the result.
* A ``Where`` with an index in the predicate still uses LINQ.
* A class that implements one of these interfaces must implement the new methods. This also
  applies to the new ``AdjustToContents`` methods below.

The ``IXLRangeColumn``, ``IXLRangeColumns``, ``IXLRangeRow`` and ``IXLRangeRows`` interfaces have
new ``AdjustToContents`` methods. They change the width of the worksheet column or the height of the
worksheet row. They use only the cells of the range. The start and end parameters are row or
column numbers of the worksheet.

**************
Data bar fill
**************

A data bar has a solid fill or a gradient fill. The new ``IXLConditionalFormat.Gradient`` property
gives the fill. The default value is ``true``, the gradient fill of Excel. Before, the save always
wrote a gradient fill, also for a solid data bar from a file.

The ``IXLCFDataBarMax`` methods ``Maximum`` and ``HighestValue`` now return the new
``IXLCFDataBar`` interface, not ``void``. Use it to set the fill:

.. code-block:: csharp

   ws.Range("A1:A5").AddConditionalFormat()
       .DataBar(XLColor.Red)
       .LowestValue()
       .HighestValue()
       .SetGradient(false);

A direct call of these methods compiles as before. Code that uses one of them as a method group for
an ``Action`` delegate does not compile. A compiled assembly that calls them must be
compiled again. A class that implements ``IXLCFDataBarMax`` or ``IXLConditionalFormat`` must
implement the changed and new members.
