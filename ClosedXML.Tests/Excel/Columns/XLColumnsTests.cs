using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Columns;

internal class XLColumnsTests
{
    [Test]
    public void Style_sets_format_of_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Columns("B-C,E").Style.Font.FontSize = 5;
        ws.Range("A1:F1").Style.Font.Bold = true; // Materialize formats in row 1

        var expected = new[]
        {
            (Column: "A", FontSize: 11),
            (Column: "B", FontSize: 5),
            (Column: "C", FontSize: 5),
            (Column: "D", FontSize: 11),
            (Column: "E", FontSize: 5),
            (Column: "F", FontSize: 11),
        };
        foreach (var (column, fontSize) in expected)
        {
            Assert.AreEqual(fontSize, ws.Column(column).Style.Font.FontSize);

            var cellWithFormat = ws.Cell(column + "1");
            Assert.AreEqual(fontSize, cellWithFormat.Style.Font.FontSize);
            Assert.True(cellWithFormat.Style.Font.Bold);

            var nonMaterializedCell = ws.Cell(column + "2");
            Assert.AreEqual(fontSize, nonMaterializedCell.Style.Font.FontSize);
        }
    }

    [Test]
    public void Enumerator_returns_columns_in_order_of_column_number()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var columns = ws.Columns("5,2,4");

        Assert.AreEqual(new[] { 2, 4, 5 }, columns.Select(c => c.ColumnNumber()).ToArray());
        var columnNumbers = new List<int>();
        foreach (IXLColumn column in (IEnumerable)columns)
            columnNumbers.Add(column.ColumnNumber());
        Assert.AreEqual(new[] { 2, 4, 5 }, columnNumbers);
    }

    [Test]
    public void AdjustToContents_uses_cells_in_row_interval()
    {
        using var wb = new XLWorkbook();
        var ws = CreateSheetWithLongText(wb);
        var columns = ws.Columns("2,4");

        Assert.AreSame(columns, columns.AdjustToContents());
        Assert.Greater(ws.Column(2).Width, 20);
        Assert.AreEqual(ws.Column(2).Width, ws.Column(4).Width);

        // The long text is in row 3. Rows outside of the interval don't change the width.
        Assert.AreSame(columns, columns.AdjustToContents(4));
        Assert.AreEqual(ws.ColumnWidth, ws.Column(2).Width);

        Assert.AreSame(columns, columns.AdjustToContents(1, 3));
        Assert.Greater(ws.Column(2).Width, 20);

        Assert.AreSame(columns, columns.AdjustToContents(1, 2));
        Assert.AreEqual(ws.ColumnWidth, ws.Column(4).Width);
    }

    [Test]
    public void AdjustToContents_keeps_width_between_limits()
    {
        // The width is a whole number of pixels, so the width can be a little different from the limit.
        using var wb = new XLWorkbook();
        var ws = CreateSheetWithLongText(wb);
        var columns = ws.Columns("2,4");

        Assert.AreSame(columns, columns.AdjustToContents(50.0, 60.0));
        Assert.AreEqual(50, ws.Column(2).Width, 0.2);

        Assert.AreSame(columns, columns.AdjustToContents(1, 55.0, 60.0));
        Assert.AreEqual(55, ws.Column(4).Width, 0.2);

        Assert.AreSame(columns, columns.AdjustToContents(1, 3, 5.0, 10.0));
        Assert.AreEqual(10, ws.Column(2).Width, 0.2);
    }

    [Test]
    public void Unhide_shows_hidden_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var columns = ws.Columns("2,4");
        columns.Hide();

        columns.Unhide();

        Assert.False(ws.Column(2).IsHidden);
        Assert.False(ws.Column(4).IsHidden);
    }

    [Test]
    public void Group_with_outline_level_sets_level_of_each_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var columns = ws.Columns("2,4");

        columns.Group(3);
        Assert.AreEqual(3, ws.Column(2).OutlineLevel);
        Assert.AreEqual(3, ws.Column(4).OutlineLevel);
        Assert.False(((XLColumn)ws.Column(2)).Collapsed);

        columns.Ungroup();
        Assert.AreEqual(2, ws.Column(4).OutlineLevel);

        columns.Group(2, true);
        Assert.AreEqual(2, ws.Column(2).OutlineLevel);
        Assert.True(((XLColumn)ws.Column(2)).Collapsed);

        columns.Collapse();
        Assert.True(ws.Column(4).IsHidden);

        columns.Expand();
        Assert.False(ws.Column(4).IsHidden);
        Assert.False(((XLColumn)ws.Column(4)).Collapsed);
    }

    [Test]
    public void Cells_returns_cells_of_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C3").Value = 2;
        ws.Cell("D4").Style.Fill.BackgroundColor = XLColor.Red;
        var columns = ws.Columns("2,4");

        Assert.AreEqual(2 * XLHelper.MaxRowNumber, columns.Cells().Count());
        Assert.AreEqual(new[] { "B2" }, columns.CellsUsed().Select(c => c.Address.ToString()).ToArray());
        Assert.AreEqual(new[] { "B2" }, columns.CellsUsed(false).Select(c => c.Address.ToString()).ToArray());
        Assert.AreEqual(new[] { "B2", "D4" }, columns.CellsUsed(true).Select(c => c.Address.ToString()).ToArray());
    }

    [Test]
    public void AddVerticalPageBreaks_adds_break_for_each_column()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var columns = ws.Columns("2,4");

        Assert.AreSame(columns, columns.AddVerticalPageBreaks());

        Assert.AreEqual(new[] { 2, 4 }, ws.PageSetup.ColumnBreaks.ToArray());
    }

    [Test]
    public void Clear_removes_content_of_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C2").Value = 2;
        ws.Cell("D2").Value = 3;
        var columns = ws.Columns("2,4");

        Assert.AreSame(columns, columns.Clear(XLClearOptions.Contents));

        Assert.AreEqual(new[] { "C2" }, ws.CellsUsed().Select(c => c.Address.ToString()).ToArray());
    }

    [Test]
    public void Delete_of_all_columns_of_sheet_removes_all_cells()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Column(3).Width = 30;

        ws.Columns().Delete();

        Assert.IsEmpty(ws.CellsUsed(XLCellsUsedOptions.All));
        Assert.AreEqual(ws.ColumnWidth, ws.Column(3).Width);
    }

    [Test]
    public void Add_to_lazy_columns_keeps_existing_columns()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = 1;
        var columns = (XLColumns)ws.Columns();

        columns.Add((XLColumn)ws.Column(5));

        Assert.AreEqual(new[] { 2, 5 }, columns.Select(c => c.ColumnNumber()).ToArray());
    }

    private static IXLWorksheet CreateSheetWithLongText(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet();
        ws.Cell("B3").Value = "A long text that needs a wide column";
        ws.Cell("D3").Value = "A long text that needs a wide column";
        return ws;
    }
}
