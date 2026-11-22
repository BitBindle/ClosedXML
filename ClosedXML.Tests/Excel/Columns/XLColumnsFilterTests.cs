using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Columns;

/// <summary>
/// The <c>Where</c>, <c>Skip</c> and <c>Take</c> methods of <see cref="IXLColumns"/> give a collection of the same type.
/// </summary>
internal class XLColumnsFilterTests
{
    private const string LongText = "A text that is much wider than the default column width";

    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Skip_then_AdjustToContents_keeps_width_of_skipped_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = LongText;
        ws.Cell("B1").Value = LongText;
        ws.Column(1).Width = 5;
        ws.Column(2).Width = 5;

        ws.ColumnsUsed().Skip(1).AdjustToContents();

        Assert.AreEqual(5, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.Greater(ws.Column(2).Width, 5);
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Where_Skip_and_Take_use_column_number_order()
    {
        // The cells are set out of order. The result has the order of the enumeration.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").Value = "c";
        ws.Cell("A1").Value = "a";
        ws.Cell("B1").Value = "b";

        CollectionAssert.AreEqual(new[] { 2, 3 }, ws.ColumnsUsed().Skip(1).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 1, 3 }, ws.ColumnsUsed().Where(c => c.ColumnNumber() != 2).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 1, 2 }, ws.ColumnsUsed().Take(2).Select(c => c.ColumnNumber()));
        CollectionAssert.AreEqual(new[] { 3, 5 }, ws.Columns("E,C,A").Skip(1).Select(c => c.ColumnNumber()));
        CollectionAssert.IsEmpty(ws.ColumnsUsed().Skip(10));
        CollectionAssert.IsEmpty(ws.ColumnsUsed().Take(-1));
        CollectionAssert.AreEqual(new[] { 1, 2, 3 }, ws.ColumnsUsed().Skip(-1).Select(c => c.ColumnNumber()));
    }

    [Test]
    public void Methods_can_be_chained()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:E1").Value = "x";

        var columns = ws.ColumnsUsed().Skip(1).Where(c => c.ColumnNumber() != 4).Take(2);

        CollectionAssert.AreEqual(new[] { 2, 3 }, columns.Select(c => c.ColumnNumber()));
    }

    [Test]
    public void Result_of_all_columns_changes_only_its_columns()
    {
        // The result of a filter is a list of columns. It does not represent all columns of the sheet.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C1").Value = "x";
        var defaultWidth = ws.ColumnWidth;

        ws.Columns().Skip(1).Width = 100;

        Assert.AreEqual(defaultWidth, ws.ColumnWidth, XLHelper.Epsilon);
        Assert.AreEqual(defaultWidth, ws.Column(1).Width, XLHelper.Epsilon);
        Assert.AreEqual(100, ws.Column(2).Width, XLHelper.Epsilon);
        Assert.AreEqual(100, ws.Column(3).Width, XLHelper.Epsilon);
        Assert.AreEqual(defaultWidth, ws.Column("Z").Width, XLHelper.Epsilon);

        ws.Columns().Where(c => c.ColumnNumber() > 1).Delete();

        Assert.AreEqual("x", ws.Cell("A1").GetText());
        Assert.IsTrue(ws.Cell("B1").IsEmpty());
        Assert.IsTrue(ws.Cell("C1").IsEmpty());
    }

    [Test]
    public void Result_keeps_columns_of_other_sheets_unchanged()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();
        ws1.Range("A1:C1").Value = "x";
        ws2.Range("A1:C1").Value = "y";

        var columns = ws1.ColumnsUsed().Take(2);
        columns.Style.Font.Bold = true;
        columns.Hide();

        Assert.IsTrue(ws1.Column(2).IsHidden);
        Assert.IsTrue(ws1.Cell("B1").Style.Font.Bold);
        Assert.IsFalse(ws1.Column(3).IsHidden);
        Assert.IsFalse(ws2.Column(2).IsHidden);
        Assert.IsFalse(ws2.Cell("B1").Style.Font.Bold);
    }

    [Test]
    public void Where_with_null_predicate_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        Assert.Throws<ArgumentNullException>(() => ws.Columns().Where((Func<IXLColumn, bool>)null!));
    }

    [Test]
    public void Where_with_index_uses_linq()
    {
        // The interface has no method for a predicate with an index, so the call uses LINQ.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C1").Value = "x";

        var columns = ws.ColumnsUsed().Where((_, index) => index > 0);

        CollectionAssert.AreEqual(new[] { 2, 3 }, columns.Select(c => c.ColumnNumber()));
    }

    [Test]
    public void FindColumns_returns_columns_of_each_sheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("One");
        var ws2 = wb.AddWorksheet("Two");
        ws1.Range("A1:C1").Value = "x";
        ws2.Range("B1:D1").Value = "y";

        var columns = wb.FindColumns(c => c.ColumnNumber() == 2);

        CollectionAssert.AreEquivalent(new[] { "One", "Two" }, columns.Select(c => c.Worksheet.Name));
    }
}
