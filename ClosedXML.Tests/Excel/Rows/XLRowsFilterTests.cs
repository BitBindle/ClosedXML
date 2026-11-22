using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Rows;

/// <summary>
/// The <c>Where</c>, <c>Skip</c> and <c>Take</c> methods of <see cref="IXLRows"/> give a collection of the same type.
/// </summary>
internal class XLRowsFilterTests
{
    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Skip_then_AdjustToContents_keeps_height_of_skipped_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "x";
        ws.Cell("A2").Value = "x";
        ws.Range("A1:A2").Style.Font.FontSize = 30;
        ws.Row(1).Height = 5;
        ws.Row(2).Height = 5;

        ws.RowsUsed().Skip(1).AdjustToContents();

        Assert.AreEqual(5, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.Greater(ws.Row(2).Height, 5);
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Where_Skip_and_Take_use_row_number_order()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A3").Value = "c";
        ws.Cell("A1").Value = "a";
        ws.Cell("A2").Value = "b";

        CollectionAssert.AreEqual(new[] { 2, 3 }, ws.RowsUsed().Skip(1).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 1, 3 }, ws.RowsUsed().Where(r => r.RowNumber() != 2).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 1, 2 }, ws.RowsUsed().Take(2).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 3, 5 }, ws.Rows("5,3,1").Skip(1).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 2 }, ws.RowsUsed().Skip(1).Where(r => r.RowNumber() != 3).Take(5).Select(r => r.RowNumber()));
    }

    [Test]
    public void Result_of_all_rows_changes_only_its_rows()
    {
        // The result of a filter is a list of rows. It does not represent all rows of the sheet.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A3").Value = "x";
        var defaultHeight = ws.RowHeight;

        ws.Rows().Skip(1).Height = 50;

        Assert.AreEqual(defaultHeight, ws.RowHeight, XLHelper.Epsilon);
        Assert.AreEqual(defaultHeight, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.AreEqual(50, ws.Row(2).Height, XLHelper.Epsilon);
        Assert.AreEqual(50, ws.Row(3).Height, XLHelper.Epsilon);
        Assert.AreEqual(defaultHeight, ws.Row(100).Height, XLHelper.Epsilon);

        ws.Rows().Where(r => r.RowNumber() > 1).Delete();

        Assert.AreEqual("x", ws.Cell("A1").GetText());
        Assert.IsTrue(ws.Cell("A2").IsEmpty());
        Assert.IsTrue(ws.Cell("A3").IsEmpty());
    }

    [Test]
    public void Where_with_null_predicate_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        Assert.Throws<ArgumentNullException>(() => ws.Rows().Where((Func<IXLRow, bool>)null!));
    }

    [Test]
    public void FindRows_returns_rows_of_each_sheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("One");
        var ws2 = wb.AddWorksheet("Two");
        ws1.Range("A1:A3").Value = "x";
        ws2.Range("A2:A4").Value = "y";

        var rows = wb.FindRows(r => r.RowNumber() == 2);

        CollectionAssert.AreEquivalent(new[] { "One", "Two" }, rows.Select(r => r.Worksheet.Name));
    }
}
