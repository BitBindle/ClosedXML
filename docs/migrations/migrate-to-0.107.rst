#############################
Migration from 0.106 to 0.107
#############################

*****************
Target frameworks
*****************

ClosedXML now targets only ``net10.0``. This version does not have the ``netstandard2.0`` and
``netstandard2.1`` builds. .NET Framework, Mono and Unity applications can't use this version.
These applications can stay on version 0.106.

*****************
Equality methods
*****************

``XLColor.Equals(object)`` now returns false for null and for an object that is not an ``XLColor``.
Before, it threw an exception. The ``IEquatable<T>.Equals`` methods of a cell fill, of a differential
number format and of a differential protection now return false for null. Before, they threw
``NullReferenceException``.

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
* ``IXLWorkbook.TryGetWorksheet`` now has the parameter
  ``[NotNullWhen(true)] out IXLWorksheet? worksheet``.
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
