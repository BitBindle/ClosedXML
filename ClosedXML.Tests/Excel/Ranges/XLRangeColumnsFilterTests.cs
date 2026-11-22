using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges;

/// <summary>
/// Tests of <c>AdjustToContents</c>, <c>Where</c>, <c>Skip</c> and <c>Take</c> for range columns.
/// </summary>
internal class XLRangeColumnsFilterTests
{
    private const string LongText = "A text that is much wider than the default column width";

    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Skip_then_AdjustToContents_keeps_width_of_skipped_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:B1").Value = LongText;
        ws.Columns(1, 2).Width = 5;

        ws.Range("A1:B1").ColumnsUsed().Skip(1).AdjustToContents();

        Assert.AreEqual(5, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 5);
    }

    [Test]
    public void AdjustToContents_of_range_column_uses_only_cells_of_range()
    {
        // The start row and the end row are row numbers of the worksheet.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = "x";
        ws.Cell("B10").Value = LongText;
        var rangeColumn = ws.Range("B1:B10").FirstColumn();

        var widthOfRange = Width(ws, () => rangeColumn.AdjustToContents());
        var widthOfFirstRow = Width(ws, () => rangeColumn.AdjustToContents(1, 1));
        var widthFromRow10 = Width(ws, () => rangeColumn.AdjustToContents(10));
        var widthWithoutRow10 = Width(ws, () => ws.Range("B1:B9").FirstColumn().AdjustToContents());

        Assert.Greater(widthOfRange, widthOfFirstRow);
        Assert.AreEqual(widthOfRange, widthFromRow10, XLHelper.Epsilon);
        Assert.AreEqual(widthOfFirstRow, widthWithoutRow10, XLHelper.Epsilon);
        Assert.AreEqual(3, Width(ws, () => rangeColumn.AdjustToContents(2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(3, Width(ws, () => rangeColumn.AdjustToContents(1, 2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(50, Width(ws, () => rangeColumn.AdjustToContents(1, 1, 50.0, 60.0)), XLHelper.Epsilon);
        Assert.AreSame(rangeColumn, rangeColumn.AdjustToContents());
    }

    [Test]
    public void AdjustToContents_of_range_columns_uses_only_cells_of_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = "x";
        ws.Cell("B10").Value = LongText;
        var rangeColumns = ws.Range("A1:B10").Columns();

        var widthOfRange = Width(ws, () => rangeColumns.AdjustToContents());
        var widthOfFirstRow = Width(ws, () => rangeColumns.AdjustToContents(1, 1));
        var widthFromRow10 = Width(ws, () => rangeColumns.AdjustToContents(10));

        Assert.Greater(widthOfRange, widthOfFirstRow);
        Assert.AreEqual(widthOfRange, widthFromRow10, XLHelper.Epsilon);
        Assert.AreEqual(3, Width(ws, () => rangeColumns.AdjustToContents(2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(3, Width(ws, () => rangeColumns.AdjustToContents(1, 2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(50, Width(ws, () => rangeColumns.AdjustToContents(1, 1, 50.0, 60.0)), XLHelper.Epsilon);
        Assert.AreSame(rangeColumns, rangeColumns.AdjustToContents());
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Where_Skip_and_Take_use_column_number_order()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("C1:E1");
        range.Value = "x";

        CollectionAssert.AreEqual(new[] { 4, 5 }, range.Columns().Skip(1).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 3, 5 }, range.Columns().Where(c => c.ColumnNumber() != 4).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 3, 4 }, range.Columns().Take(2).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 4 }, range.Columns().Skip(1).Take(1).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 4, 5 }, range.Columns("3,2").Take(5).Select(c => c.ColumnNumber()));
    }

    [Test]
    public void Skip_then_Delete_keeps_first_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "keep";
        ws.Range("B1:C1").Value = "remove";

        ws.Range("A1:C1").Columns().Skip(1).Delete();

        Assert.AreEqual("keep", ws.Cell("A1").GetText());
        Assert.IsTrue(ws.Cell("B1").IsEmpty());
        Assert.IsTrue(ws.Cell("C1").IsEmpty());
    }

    [Test]
    public void Where_with_null_predicate_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        Assert.Throws<ArgumentNullException>(() => ws.Range("A1:C1").Columns().Where(null!));
    }

    private static double Width(IXLWorksheet ws, Action adjust)
    {
        ws.Column(2).Width = 1;
        adjust();
        return ws.Column(2).Width;
    }
}
