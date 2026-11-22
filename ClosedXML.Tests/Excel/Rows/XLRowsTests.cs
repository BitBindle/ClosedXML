using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Rows;

internal class XLRowsTests
{
    [Test]
    public void Style_sets_format_of_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Rows("2,4-5").Style.Font.FontSize = 5;
        ws.Range("A1:A6").Style.Font.Bold = true; // Materialize formats in column A

        var expected = new[]
        {
            (Row: 1, FontSize: 11),
            (Row: 2, FontSize: 5),
            (Row: 3, FontSize: 11),
            (Row: 4, FontSize: 5),
            (Row: 5, FontSize: 5),
            (Row: 6, FontSize: 11),
        };
        foreach (var (row, fontSize) in expected)
        {
            Assert.AreEqual(fontSize, ws.Row(row).Style.Font.FontSize);

            var cellWithFormat = ws.Cell("A" + row);
            Assert.AreEqual(fontSize, cellWithFormat.Style.Font.FontSize);
            Assert.True(cellWithFormat.Style.Font.Bold);

            var nonMaterializedCell = ws.Cell("B" + row);
            Assert.AreEqual(fontSize, nonMaterializedCell.Style.Font.FontSize);
        }
    }

    [Test]
    public void Enumerator_returns_rows_in_order_of_row_number()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var rows = ws.Rows("5,2,4");

        Assert.AreEqual(new[] { 2, 4, 5 }, rows.Select(r => r.RowNumber()).ToArray());
        var rowNumbers = new List<int>();
        foreach (IXLRow row in (IEnumerable)rows)
            rowNumbers.Add(row.RowNumber());
        Assert.AreEqual(new[] { 2, 4, 5 }, rowNumbers);
    }

    [Test]
    public void AdjustToContents_uses_cells_in_column_interval()
    {
        using var wb = new XLWorkbook();
        var ws = CreateSheetWithLargeFont(wb);
        var rows = ws.Rows("2,4");

        Assert.AreSame(rows, rows.AdjustToContents());
        Assert.Greater(ws.Row(2).Height, 30);
        Assert.AreEqual(ws.Row(2).Height, ws.Row(4).Height);

        // The large font is in column C. Columns outside of the interval don't change the height.
        Assert.AreSame(rows, rows.AdjustToContents(4));
        Assert.AreEqual(ws.RowHeight, ws.Row(2).Height);

        Assert.AreSame(rows, rows.AdjustToContents(1, 3));
        Assert.Greater(ws.Row(2).Height, 30);

        Assert.AreSame(rows, rows.AdjustToContents(1, 2));
        Assert.AreEqual(ws.RowHeight, ws.Row(4).Height);
    }

    [Test]
    public void AdjustToContents_keeps_height_between_limits()
    {
        using var wb = new XLWorkbook();
        var ws = CreateSheetWithLargeFont(wb);
        var rows = ws.Rows("2,4");

        Assert.AreSame(rows, rows.AdjustToContents(50.0, 60.0));
        Assert.AreEqual(50, ws.Row(2).Height);

        Assert.AreSame(rows, rows.AdjustToContents(1, 55.0, 60.0));
        Assert.AreEqual(55, ws.Row(4).Height);

        Assert.AreSame(rows, rows.AdjustToContents(1, 3, 10.0, 20.0));
        Assert.AreEqual(20, ws.Row(2).Height);
    }

    [Test]
    public void Unhide_shows_hidden_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var rows = ws.Rows("2,4");
        rows.Hide();

        rows.Unhide();

        Assert.False(ws.Row(2).IsHidden);
        Assert.False(ws.Row(4).IsHidden);
    }

    [Test]
    public void Group_with_outline_level_sets_level_of_each_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var rows = ws.Rows("2,4");

        rows.Group(3);
        Assert.AreEqual(3, ws.Row(2).OutlineLevel);
        Assert.AreEqual(3, ws.Row(4).OutlineLevel);
        Assert.False(((XLRow)ws.Row(2)).Collapsed);

        rows.Ungroup();
        Assert.AreEqual(2, ws.Row(4).OutlineLevel);

        rows.Group(2, true);
        Assert.AreEqual(2, ws.Row(2).OutlineLevel);
        Assert.True(((XLRow)ws.Row(2)).Collapsed);

        rows.Collapse();
        Assert.True(ws.Row(4).IsHidden);

        rows.Expand();
        Assert.False(ws.Row(4).IsHidden);
        Assert.False(((XLRow)ws.Row(4)).Collapsed);
    }

    [Test]
    public void Cells_returns_cells_of_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C3").Value = 2;
        ws.Cell("D4").Style.Fill.BackgroundColor = XLColor.Red;
        var rows = ws.Rows("2,4");

        Assert.AreEqual(2 * XLHelper.MaxColumnNumber, rows.Cells().Count());
        Assert.AreEqual(new[] { "B2" }, rows.CellsUsed().Select(c => c.Address.ToString()).ToArray());
        Assert.AreEqual(new[] { "B2", "D4" }, rows.CellsUsed(XLCellsUsedOptions.All).Select(c => c.Address.ToString()).ToArray());
    }

    [Test]
    public void AddHorizontalPageBreaks_adds_break_for_each_row()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var rows = ws.Rows("2,4");

        Assert.AreSame(rows, rows.AddHorizontalPageBreaks());

        Assert.AreEqual(new[] { 2, 4 }, ws.PageSetup.RowBreaks.ToArray());
    }

    [Test]
    public void Clear_removes_content_of_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("B3").Value = 2;
        ws.Cell("B4").Value = 3;
        var rows = ws.Rows("2,4");

        Assert.AreSame(rows, rows.Clear(XLClearOptions.Contents));

        Assert.AreEqual(new[] { "B3" }, ws.CellsUsed().Select(c => c.Address.ToString()).ToArray());
    }

    [Test]
    public void Delete_of_all_rows_of_sheet_removes_all_cells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Row(3).Height = 30;

        ws.Rows().Delete();

        Assert.IsEmpty(ws.CellsUsed(XLCellsUsedOptions.All));
        Assert.AreEqual(ws.RowHeight, ws.Row(3).Height);
    }

    [Test]
    public void Add_to_lazy_rows_keeps_existing_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A2").Value = 1;
        var rows = (XLRows)ws.Rows();

        rows.Add((XLRow)ws.Row(5));

        Assert.AreEqual(new[] { 2, 5 }, rows.Select(r => r.RowNumber()).ToArray());
    }

    private static IXLWorksheet CreateSheetWithLargeFont(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet();
        ws.Cell("C2").SetValue("Text").Style.Font.FontSize = 30;
        ws.Cell("C4").SetValue("Text").Style.Font.FontSize = 30;
        return ws;
    }
}
