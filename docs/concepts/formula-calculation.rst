*******************
Formula calculation
*******************

Each cell in a sheet contains a value. For most cells, the value is
``Blank.Value``. Sheet can also contain formulas that determine values of
other cells during recalculation.

Formula can modify one of more cells, normal formula are only for one cell,
array or data table formula can change many cells.

Please do read the `Excel Recalculation <https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation>`_
artible by Microsoft that describes the fundamentals.

Cell values
===========

``IXLCell`` provides two properties that give access to value of a cell:
``IXLCell.Value`` and ``IXLCell.CachedValue``.

There is an important difference between them:

* ``IXLCell.Value`` will check if cell contains a dirty formula and if it
  does, **the cell value is recalculated** before the value is returned.
  The getter always returns up-to-date value.
* ``IXLCell.CachedValue`` returns a value of a cell, even if the value is
  stale.

Using the ``IXLCell.Value`` thus carries a performance impact and if possible, the
``IXLCell.CachedValue`` should be used instead.

In order to acheve a performant code while working with formulas

* Use ``IXLCell.InsertData`` to insert data. It marks all formulas dependent
  on the inserted cells as dirty.
* Use ``IXLCell.Value`` setter for setting a single cell value. Setting a
  single value also triggers dirty marking (among other things), but only
  for the single cell. It is better to use ``IXLCell.InsertData`` for larger
  amount of data.
* Use ``IXLCell.CachedValue`` getter to retrieve a value of a cell.
* Use ``SaveOptions.EvaluateFormulasBeforeSaving`` to recalculate dirty
  formulas before saving the data to a file.
* Use ``IXLWorkbook.RecalculateAllFormulas()`` or
  ``IXLWorksheet.RecalculateAllFormulas()`` to update formulas. Note that this
  recalculates all formulas, not just non-dirty ones.
* Use ``IXLWorkbook.RecalculateDirtyFormulas()`` to recalculate only the dirty
  formulas, e.g. after a change of a few cells. The overload with a collection
  of failed cells doesn't stop at a cell that fails.

Dirty tracking
==============

Each formula can be dirty or non-dirty. Formula is dirty, if the value is
calculated last time might no longer be accurate. That can happen for several
reasons:

* Another cell the formula depends on changed value (e.g. ``=A1+2`` and ``A1``
  changed value from ``1`` to ``2``).
* Worksheet has been added or removed and thus some values might be available
  or are no longer available (e.g. ``=Test!A1``, but sheet ``Test`` didn't
  exist previously and was just added).
* Defined name was added or removed

Whether a formula is dirty can be checked through ``IXLCell.NeedsRecalculation``
property. Cells without formulas are never dirty.

Use ``IXLCell.InvalidateFormula()`` to mark the formula as dirty.

Some changes don't set a value of a cell, so they don't mark a formula as
dirty. Examples are a hide of a row for ``SUBTOTAL(109, ...)`` and a new time
for ``NOW()``. Use ``IXLWorkbook.MarkDependentsDirty(range)`` to mark the
formulas that read the range as dirty:

.. code-block:: csharp

   ws.Row(2).Hide();
   wb.MarkDependentsDirty(ws.Row(2));
   wb.RecalculateDirtyFormulas();

The mark doesn't go past a formula that is already dirty. After
``IXLCell.InvalidateFormula()``, call ``MarkDependentsDirty(cell.AsRange())``
so that the formulas that read the cell become dirty too.

A set of a value marks the formulas that read the cell only when the
dependency tree exists. A loaded workbook has no tree. Each calculation builds
it, ``RecalculateDirtyFormulas()`` included, and so does
``MarkDependentsDirty``. Call ``RecalculateDirtyFormulas()`` once after a load
and before the first change.

Transitive nature
-----------------

ClosedXML holds information about what areas formula depends on in a
`r-tree <https://en.wikipedia.org/wiki/R-tree>`_ (e.g. ``=SUM(A1:D4) + C7``
depends on ``A1:D4`` and ``C7:C7``). When an area of a sheet changes (e.g.
``B2`` is now ``7`` instead of ``5``), the r-tree finds all overlapping
areas that formulas depends on (e.g. ``A1:D4`` intersects with ``B2``) and
marks all those formulas as dirty.

This process continues, because cell of the formula is now also dirty and that
can affect other formulas.

Loaded formulas
---------------

When a formula is loaded and the cell also contains a value, the formula is not
marked as dirty on load. That is done under assumption that file contains the
truth about what value of a cell should be.

If a cell contains a formula, but not a value, the formula is marked as a dirty
during load.

Time of the calculation
=======================

``TODAY`` and ``NOW`` return the local time of the clock of the workbook. The
conversion of a date text without a year (e.g. ``"5-Jul"``) uses the year of
this time, and so does the century of a two-digit year (e.g. ``"Mar-35"``).
The default clock is ``TimeProvider.System``. Set ``LoadOptions.TimeProvider``
to use a fixed time or the time zone of a user:

.. code-block:: csharp

   var options = new LoadOptions { TimeProvider = userClock };
   using var wb = new XLWorkbook(options);

``TimeProvider.GetLocalNow()`` gives the local time in the
``TimeProvider.LocalTimeZone``. A change of the time doesn't mark a formula as
dirty. Call ``RecalculateAllFormulas()`` to calculate the formulas again.
A conversion outside of a formula (e.g. ``XLCellValue.TryConvert``) uses the
system clock.

Calculation chain
=================

Calculation chain contains all formula cells. The cells in the chain are in
the order in which they were evalualted during last formula calculation. Cell
should be in order of dependencies (e.g. if ``A2`` depends on ``A1``, ``A1``
should be positioned before the ``A2`` in the chain).

.. image:: img/calc-chain-formulas.png 
  :alt: Formulas in a sheet with precedent arrows

The "correct" order of the chain from the image might be any of following:

* B1, C1, D1, B4, C4
* B1, B4, C1, C4, D1
* B1, C1, B4, C4, D1
* B1, C1, B4, D1, C4 
* ...

The chain does't have to have correct order, it has the correct order during
last calculation (or any order, if no calculation was yet done). In the
beginning, the chain could look like this: D1, C1, B1, B4, C4.

When calc engine calculates formulas, it goes through the chain and tries to
evaluate each formula. If the formula depends on a cell that is dirty, it
stops evaluation of the current formula and moves the supporting formula before
the current one and start to evaluate the new current formula.

This way, chain should end up with a correct order.

.. note::
   The order of cells changes most during the first calculation and far less in
   subsequent calculations. That is the reason why first calculation is
   generally the slowest.

Cyclic dependencies
-------------------

Calc engine will throw an ``InvalidOperationException``, when it encounteres a
cycle in a calculation chain.

Blocked names
=============

A workbook from an unknown source can have defined names with formulas that
the application doesn't trust, or tables that it doesn't support. Set
``LoadOptions.CalculationNameFilter`` to block defined names and tables. The
function gets the sheet name of a sheet-scoped name, or ``null`` for a
workbook-scoped name, a table or a name that the workbook doesn't have, and the
name. It returns ``false`` to block the name:

.. code-block:: csharp

   var trusted = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Rate" };
   var options = new LoadOptions
   {
       CalculationNameFilter = (sheetName, name) => sheetName is null && trusted.Contains(name)
   };
   using var wb = new XLWorkbook(stream, options);

The function gets the stored name of a defined name or a table. For a name that
the workbook doesn't have, it gets the text of the formula, so compare names
without case (``StringComparer.OrdinalIgnoreCase`` above).

A formula that uses a blocked name throws an ``InvalidOperationException``,
also through ``INDIRECT`` or a structured reference.
``RecalculateAllFormulas(failed)`` adds the cell to the failed cells. The
calculation doesn't read the formula of a blocked name. A sheet-scoped name
hides the workbook-scoped name with the same name on its sheet, also when the
function blocks the sheet-scoped name.

Records of a recalculation
==========================

The result of some operations depends on the locale of the application that
calculates the formula, e.g. ``="1.5"*2`` or ``=TEXT(A1,"0.0")``. For some
other operations, ClosedXML can give another result than Excel. Use the
overload with records to find the formula cells that use such an operation:

.. code-block:: csharp

   var failed = new List<IXLCell>();
   var records = new List<XLCalculationRecord>();
   wb.RecalculateAllFormulas(failed, records);
   foreach (var (cell, kind) in records)
       Console.WriteLine($"{cell.Address}: {kind}");

``XLCalculationRecordKind.LocaleOperation`` marks a conversion of a text with a
digit to a number or a date, a conversion of a number or a logical value to a
text, a text read as a logical value, a criterion with the name of a logical or
an error value in English or German, a text criterion on a logical or an error
value, an order comparison of two texts, an equality of two texts that both
have ``i``, ``I``, ``ı`` or ``İ``, ``TEXT``, ``FIXED`` and ``DOLLAR`` of a
number, and ``UPPER``, ``LOWER`` and ``PROPER`` of a text with one of these
letters.

``XLCalculationRecordKind.EngineDifference`` marks a ``MATCH`` lookup text or a
criterion longer than 255 characters, and a wildcard pattern with a ``~`` that
is not followed by ``*``, ``?`` or ``~``.

A cell gets at most one record of each kind. The records don't change the
values. ``RecalculateDirtyFormulas(failed, records)`` gives the records of the
dirty formulas that it calculates. The other overloads don't collect records.

Wrong number of arguments
=========================

Excel doesn't accept a formula that calls a function with too few or too many
arguments, e.g. ``=YEAR()`` or ``=NOW(1)``. A file from another application can
still contain such a formula. By default, the parse of the formula throws an
``ExpressionParseException``, so the cell isn't calculated. A read of a formula
value that starts a recalculation throws too.
``RecalculateAllFormulas(failed)`` adds the cell to the failed cells.

Set ``LoadOptions.WrongArgumentCountAsErrorValue`` to ``true`` to calculate
such a formula as ``#VALUE!``. LibreOffice writes ``#VALUE!`` to an xlsx file
for a formula that it refuses for this reason:

.. code-block:: csharp

   using var wb = new XLWorkbook(stream, new LoadOptions { WrongArgumentCountAsErrorValue = true });
   wb.Evaluate("YEAR()");            // #VALUE!
   wb.Evaluate("IFERROR(NOW(1),0)"); // #VALUE!

The whole formula is ``#VALUE!``, so a function in the formula, e.g.
``IFERROR``, doesn't change the result. A formula in another cell that reads
the cell gets ``#VALUE!``. The numbers of arguments are the numbers of Excel.
LibreOffice accepts some calls that Excel refuses, e.g. ``ROUND(1)``. A syntax
error is a parse error with each option value. A call of a function that
ClosedXML doesn't know is ``#NAME?`` with each option value, if the formula has
no wrong number of arguments.
