using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges;

[TestFixture]
internal class XLRangeColumnTests
{
    [Test]
    public void InsertCellsAbove_inserts_cells_above_column()
    {
        using var wb = new XLWorkbook();
        var column = CreateColumn(wb);
        var ws = column.Worksheet;

        var cells = column.InsertCellsAbove(2);

        Assert.AreEqual(new[] { "B2", "B3" }, GetAddresses(cells));
        Assert.AreEqual(1, ws.Cell("B4").Value);
        Assert.AreEqual(4, ws.Cell("B7").Value);
        Assert.AreEqual("B4:B7", column.RangeAddress.ToString());
    }

    [Test]
    public void InsertCellsBelow_inserts_cells_below_column()
    {
        using var wb = new XLWorkbook();
        var column = CreateColumn(wb);
        var ws = column.Worksheet;
        ws.Cell("B6").Value = "Below";

        var cells = column.InsertCellsBelow(2);

        Assert.AreEqual(new[] { "B6", "B7" }, GetAddresses(cells));
        Assert.AreEqual("Below", ws.Cell("B8").Value);
        Assert.AreEqual("B2:B7", column.RangeAddress.ToString());
    }

    [Test]
    public void CopyTo_cell_returns_column_at_target()
    {
        using var wb = new XLWorkbook();
        var column = CreateColumn(wb);
        var ws = column.Worksheet;

        var copy = column.CopyTo(ws.Cell("E3"));

        Assert.AreEqual("E3:E6", copy.RangeAddress.ToString());
        Assert.AreEqual(new XLCellValue[] { 1, 2, 3, 4 }, copy.Cells().Select(c => c.Value).ToArray());
    }

    [Test]
    public void CopyTo_range_returns_column_at_first_cell_of_target()
    {
        using var wb = new XLWorkbook();
        var column = CreateColumn(wb);
        var ws = column.Worksheet;

        var copy = column.CopyTo(ws.Range("E3:G4"));

        Assert.AreEqual("E3:E6", copy.RangeAddress.ToString());
        Assert.AreEqual(new XLCellValue[] { 1, 2, 3, 4 }, copy.Cells().Select(c => c.Value).ToArray());
    }

    [TestCase("1", new[] { "B2:B2" })]
    [TestCase("1:2", new[] { "B2:B3" })]
    [TestCase("2-3, 4", new[] { "B3:B4", "B5:B5" })]
    [TestCase("B3:B4", new[] { "B3:B4" })]
    public void Columns_returns_parts_of_column(string rows, string[] expectedAddresses)
    {
        // A number is a row number relative to the column. Other text is a cell address.
        using var wb = new XLWorkbook();
        var column = CreateColumn(wb);

        var columns = column.Columns(rows);

        Assert.AreEqual(expectedAddresses, columns.Select(c => c.RangeAddress.ToString()).ToArray());
    }

    [Test]
    public void AsTable_makes_table_that_is_not_in_worksheet_tables()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "First";
        ws.Cell("B3").Value = 1;
        ws.Cell("C2").Value = "Second";
        ws.Cell("C3").Value = 2;

        var unnamedTable = ws.Range("B2:B3").Column(1).AsTable();
        var namedTable = ws.Range("C2:C3").Column(1).AsTable("Named");

        Assert.AreEqual("B2:B3", unnamedTable.RangeAddress.ToString());
        Assert.AreEqual("C2:C3", namedTable.RangeAddress.ToString());
        Assert.AreEqual("Named", namedTable.Name);
        Assert.IsEmpty(ws.Tables);
    }

    [Test]
    public void CreateTable_with_name_adds_table_to_worksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Header";
        ws.Cell("B3").Value = 1;
        var column = ws.Range("B2:B3").Column(1);

        var table = column.CreateTable("Created");

        Assert.AreEqual("B2:B3", table.RangeAddress.ToString());
        Assert.AreEqual("Created", ws.Tables.Single().Name);
    }

    [Test]
    public void CreateTable_without_name_adds_table_to_worksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Header";
        ws.Cell("B3").Value = 1;
        var column = ws.Range("B2:B3").Column(1);

        var table = column.CreateTable();

        Assert.AreSame(table, ws.Tables.Single());
    }

    [Test]
    public void Column_of_table_cannot_become_another_table()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Header";
        ws.Cell("B3").Value = 1;
        var table = ws.Range("B2:B3").CreateTable();
        var tableColumn = table.Column(1);

        Assert.IsTrue(((XLRangeColumn)tableColumn).IsTableColumn());
        Assert.Throws<InvalidOperationException>(() => tableColumn.AsTable());
        Assert.Throws<InvalidOperationException>(() => tableColumn.AsTable("Other"));
        Assert.Throws<InvalidOperationException>(() => tableColumn.CreateTable());
        Assert.Throws<InvalidOperationException>(() => tableColumn.CreateTable("Other"));
    }

    [Test]
    public void Delete_of_table_column_needs_text_in_top_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Header";
        ws.Cell("B3").Value = 1;
        var table = ws.Range("B2:B3").CreateTable();
        table.ShowHeaderRow = false;
        var tableColumn = table.Column(1);

        // Without the header row, the top cell of the table column contains a number.
        Assert.Throws<InvalidOperationException>(() => tableColumn.Delete());
    }

    [Test]
    public void Delete_of_table_column_needs_field_name_in_top_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Header";
        ws.Cell("B3").Value = "Text";
        var table = ws.Range("B2:B3").CreateTable();
        table.ShowHeaderRow = false;
        var tableColumn = table.Column(1);

        // Without the header row, the top cell of the table column contains a text that isn't a field name.
        Assert.Throws<InvalidOperationException>(() => tableColumn.Delete());
    }

    private static IXLRangeColumn CreateColumn(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("B3").Value = 2;
        ws.Cell("B4").Value = 3;
        ws.Cell("B5").Value = 4;
        return ws.Range("B2:C5").Column(1);
    }

    private static string[] GetAddresses(IXLCells cells)
    {
        return cells.Select(c => c.Address.ToString()).ToArray();
    }
}
