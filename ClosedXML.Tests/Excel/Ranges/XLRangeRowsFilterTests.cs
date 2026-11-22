using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges;

/// <summary>
/// Tests of <c>AdjustToContents</c>, <c>Where</c>, <c>Skip</c> and <c>Take</c> for range rows.
/// </summary>
internal class XLRangeRowsFilterTests
{
    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Skip_then_AdjustToContents_keeps_height_of_skipped_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A1:A2");
        range.Value = "x";
        range.Style.Font.FontSize = 30;
        ws.Rows(1, 2).Height = 5;

        range.Rows().Skip(1).AdjustToContents();

        Assert.AreEqual(5, ws.Row(1).Height, XLHelper.Epsilon);
        Assert.Greater(ws.Row(2).Height, 5);
    }

    [Test]
    public void AdjustToContents_of_range_row_uses_only_cells_of_range()
    {
        // The start column and the end column are column numbers of the worksheet.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A2").Value = "x";
        ws.Cell("J2").Value = "x";
        ws.Cell("J2").Style.Font.FontSize = 30;
        var rangeRow = ws.Range("A2:J2").FirstRow();

        var heightOfRange = Height(ws, () => rangeRow.AdjustToContents());
        var heightOfFirstColumn = Height(ws, () => rangeRow.AdjustToContents(1, 1));
        var heightFromColumn10 = Height(ws, () => rangeRow.AdjustToContents(10));
        var heightWithoutColumn10 = Height(ws, () => ws.Range("A2:I2").FirstRow().AdjustToContents());

        Assert.Greater(heightOfRange, heightOfFirstColumn);
        Assert.AreEqual(heightOfRange, heightFromColumn10, XLHelper.Epsilon);
        Assert.AreEqual(heightOfFirstColumn, heightWithoutColumn10, XLHelper.Epsilon);
        Assert.AreEqual(3, Height(ws, () => rangeRow.AdjustToContents(2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(3, Height(ws, () => rangeRow.AdjustToContents(1, 2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(80, Height(ws, () => rangeRow.AdjustToContents(1, 1, 80.0, 90.0)), XLHelper.Epsilon);
        Assert.AreSame(rangeRow, rangeRow.AdjustToContents());
    }

    [Test]
    public void AdjustToContents_of_range_rows_uses_only_cells_of_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A2").Value = "x";
        ws.Cell("J2").Value = "x";
        ws.Cell("J2").Style.Font.FontSize = 30;
        var rangeRows = ws.Range("A1:J2").Rows();

        var heightOfRange = Height(ws, () => rangeRows.AdjustToContents());
        var heightOfFirstColumn = Height(ws, () => rangeRows.AdjustToContents(1, 1));
        var heightFromColumn10 = Height(ws, () => rangeRows.AdjustToContents(10));

        Assert.Greater(heightOfRange, heightOfFirstColumn);
        Assert.AreEqual(heightOfRange, heightFromColumn10, XLHelper.Epsilon);
        Assert.AreEqual(3, Height(ws, () => rangeRows.AdjustToContents(2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(3, Height(ws, () => rangeRows.AdjustToContents(1, 2.0, 3.0)), XLHelper.Epsilon);
        Assert.AreEqual(80, Height(ws, () => rangeRows.AdjustToContents(1, 1, 80.0, 90.0)), XLHelper.Epsilon);
        Assert.AreSame(rangeRows, rangeRows.AdjustToContents());
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#2867")]
    public void Where_Skip_and_Take_use_row_number_order()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range("A3:A5");
        range.Value = "x";

        CollectionAssert.AreEqual(new[] { 4, 5 }, range.Rows().Skip(1).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 3, 5 }, range.Rows().Where(r => r.RowNumber() != 4).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 3, 4 }, range.Rows().Take(2).Select(r => r.RowNumber()));
        CollectionAssert.AreEqual(new[] { 4, 5 }, range.Rows("3,2").Take(5).Select(r => r.RowNumber()));
    }

    [Test]
    public void Skip_then_Delete_keeps_first_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "keep";
        ws.Range("A2:A3").Value = "remove";

        ws.Range("A1:A3").Rows().Skip(1).Delete();

        Assert.AreEqual("keep", ws.Cell("A1").GetText());
        Assert.IsTrue(ws.Cell("A2").IsEmpty());
        Assert.IsTrue(ws.Cell("A3").IsEmpty());
    }

    [Test]
    public void Where_with_null_predicate_throws()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        Assert.Throws<ArgumentNullException>(() => ws.Range("A1:A3").Rows().Where(null!));
    }

    private static double Height(IXLWorksheet ws, Action adjust)
    {
        ws.Row(2).Height = 1;
        adjust();
        return ws.Row(2).Height;
    }
}
