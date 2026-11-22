using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Ranges;

[TestFixture]
internal class XLRangeRowTests
{
    [Test]
    public void Cell_with_column_letter_is_relative_to_row()
    {
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);

        var cell = row.Cell("B");

        Assert.AreEqual("C2", cell.Address.ToString());
        Assert.AreEqual(2, cell.Value);
    }

    [Test]
    public void Cells_with_column_letters_returns_cells_in_interval()
    {
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);

        var cells = row.Cells("B", "C");

        Assert.AreEqual(new[] { "C2", "D2" }, GetAddresses(cells));
    }

    [Test]
    public void CellCount_returns_number_of_columns()
    {
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);

        Assert.AreEqual(4, row.CellCount());
    }

    [Test]
    public void InsertCellsAfter_inserts_cells_right_of_row()
    {
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);
        var ws = row.Worksheet;
        ws.Cell("F2").Value = "Right";

        var cells = row.InsertCellsAfter(2);

        Assert.AreEqual(new[] { "F2", "G2" }, GetAddresses(cells));
        Assert.AreEqual("Right", ws.Cell("H2").Value);
        Assert.AreEqual("B2:G2", row.RangeAddress.ToString());
    }

    [Test]
    public void InsertCellsBefore_inserts_cells_left_of_row()
    {
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);
        var ws = row.Worksheet;

        var cells = row.InsertCellsBefore(2);

        Assert.AreEqual(new[] { "B2", "C2" }, GetAddresses(cells));
        Assert.AreEqual(1, ws.Cell("D2").Value);
        Assert.AreEqual(4, ws.Cell("G2").Value);
        Assert.AreEqual("D2:G2", row.RangeAddress.ToString());
    }

    [Test]
    public void CopyTo_cell_returns_row_at_target()
    {
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);
        var ws = row.Worksheet;

        var copy = row.CopyTo(ws.Cell("C10"));

        Assert.AreEqual("C10:F10", copy.RangeAddress.ToString());
        Assert.AreEqual(new XLCellValue[] { 1, 2, 3, 4 }, copy.Cells().Select(c => c.Value).ToArray());
    }

    [Test]
    public void CopyTo_range_returns_row_at_first_cell_of_target()
    {
        // The copy starts at the first cell of the target range. The returned row must contain the copied cells.
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);
        var ws = row.Worksheet;

        var copy = row.CopyTo(ws.Range("C10:E12"));

        Assert.AreEqual("C10:F10", copy.RangeAddress.ToString());
        Assert.AreEqual(new XLCellValue[] { 1, 2, 3, 4 }, copy.Cells().Select(c => c.Value).ToArray());
    }

    [TestCase("1", new[] { "B2:B2" })]
    [TestCase("1:2", new[] { "B2:C2" })]
    [TestCase("2-3, 4", new[] { "C2:D2", "E2:E2" })]
    [TestCase("C2:D2", new[] { "C2:D2" })]
    public void Rows_returns_parts_of_row(string columns, string[] expectedAddresses)
    {
        // A number is a column number relative to the row. Other text is a cell address.
        using var wb = new XLWorkbook();
        var row = CreateRow(wb);

        var rows = row.Rows(columns);

        Assert.AreEqual(expectedAddresses, rows.Select(r => r.RangeAddress.ToString()).ToArray());
    }

    private static IXLRangeRow CreateRow(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C2").Value = 2;
        ws.Cell("D2").Value = 3;
        ws.Cell("E2").Value = 4;
        return ws.Range("B2:E3").Row(1);
    }

    private static string[] GetAddresses(IXLCells cells)
    {
        return cells.Select(c => c.Address.ToString()).ToArray();
    }
}
