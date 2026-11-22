using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Tables;

[TestFixture]
internal class XLTableRangeTests
{
    [Test]
    public void FirstRow_without_predicate_returns_first_data_row()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var row = dataRange.FirstRow();

        Assert.AreEqual("B3:C3", row.RangeAddress.ToString());
    }

    [Test]
    public void FirstRow_with_predicate_returns_first_matched_row()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var emptyRow = dataRange.FirstRow(r => r.Field("Name").IsEmpty());
        var noRow = dataRange.FirstRow(_ => false);

        Assert.AreEqual("B4:C4", emptyRow.RangeAddress.ToString());
        Assert.IsNull(noRow);
    }

    [Test]
    public void FirstRowUsed_skips_empty_rows()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        Assert.AreEqual("B3:C3", dataRange.FirstRowUsed().RangeAddress.ToString());
        Assert.AreEqual("B5:C5", dataRange.FirstRowUsed(r => r.Field("Name").GetString() != "Cake").RangeAddress.ToString());
        Assert.AreEqual("B6:C6", dataRange.FirstRowUsed(XLCellsUsedOptions.AllContents, r => r.Field("Name").GetString() == "Pie").RangeAddress.ToString());
        Assert.AreEqual("B3:C3", dataRange.FirstRowUsed(XLCellsUsedOptions.AllContents).RangeAddress.ToString());
        Assert.IsNull(dataRange.FirstRowUsed(_ => false));
    }

    [Test]
    public void LastRow_without_predicate_returns_last_data_row()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var row = dataRange.LastRow();

        Assert.AreEqual("B7:C7", row.RangeAddress.ToString());
    }

    [Test]
    public void LastRow_with_predicate_returns_last_matched_row()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var row = dataRange.LastRow(r => !r.Field("Name").IsEmpty());
        var noRow = dataRange.LastRow(_ => false);

        Assert.AreEqual("B6:C6", row.RangeAddress.ToString());
        Assert.IsNull(noRow);
    }

    [Test]
    public void LastRowUsed_skips_empty_rows()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        Assert.AreEqual("B6:C6", dataRange.LastRowUsed().RangeAddress.ToString());
        Assert.AreEqual("B5:C5", dataRange.LastRowUsed(r => r.Field("Name").GetString() != "Pie").RangeAddress.ToString());
        Assert.AreEqual("B3:C3", dataRange.LastRowUsed(XLCellsUsedOptions.AllContents, r => r.Field("Name").GetString() == "Cake").RangeAddress.ToString());
        Assert.AreEqual("B6:C6", dataRange.LastRowUsed(XLCellsUsedOptions.AllContents).RangeAddress.ToString());
        Assert.IsNull(dataRange.LastRowUsed(_ => false));
    }

    [Test]
    public void Row_number_must_be_positive()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        Assert.Throws<ArgumentOutOfRangeException>(() => dataRange.Row(0));
    }

    [Test]
    public void Row_returns_table_row_with_access_to_fields()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var row = dataRange.Row(3);

        Assert.AreEqual("B5:C5", row.RangeAddress.ToString());
        Assert.AreEqual("Waffle", row.Field("Name").GetString());
    }

    [Test]
    public void Rows_with_predicate_returns_matched_rows()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var allRows = dataRange.Rows();
        var rowsWithQuantity = dataRange.Rows(r => r.Field("Qty").Value.IsNumber);

        Assert.AreEqual(new[] { "B3:C3", "B4:C4", "B5:C5", "B6:C6", "B7:C7" }, GetAddresses(allRows));
        Assert.AreEqual(new[] { "B3:C3", "B5:C5", "B6:C6" }, GetAddresses(rowsWithQuantity));
    }

    [Test]
    public void Rows_with_numbers_returns_rows_in_interval()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var rows = dataRange.Rows(2, 4);

        Assert.AreEqual(new[] { "B4:C4", "B5:C5", "B6:C6" }, GetAddresses(rows));
    }

    [TestCase("1", new[] { "B3:C3" })]
    [TestCase("1, 3", new[] { "B3:C3", "B5:C5" })]
    [TestCase("2:3", new[] { "B4:C4", "B5:C5" })]
    [TestCase("1,3-5", new[] { "B3:C3", "B5:C5", "B6:C6", "B7:C7" })]
    public void Rows_with_text_returns_listed_rows(string rows, string[] expectedAddresses)
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var tableRows = dataRange.Rows(rows);

        Assert.AreEqual(expectedAddresses, GetAddresses(tableRows));
    }

    [Test]
    public void RowsUsed_skips_empty_rows()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var usedRows = dataRange.RowsUsed();
        var matchedRows = dataRange.RowsUsed(r => r.Field("Qty").GetDouble() > 5);
        var rowsWithOptions = dataRange.RowsUsed(XLCellsUsedOptions.AllContents, r => r.Field("Name").GetString() == "Waffle");

        Assert.AreEqual(new[] { "B3:C3", "B5:C5", "B6:C6" }, GetAddresses(usedRows));
        Assert.AreEqual(new[] { "B3:C3", "B6:C6" }, GetAddresses(matchedRows));
        Assert.AreEqual(new[] { "B5:C5" }, GetAddresses(rowsWithOptions));
    }

    [Test]
    public void Table_returns_table_of_data_range()
    {
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        Assert.AreEqual("Products", dataRange.Table.Name);
    }

    [TestCase("A", "B3:B7")]
    [TestCase("B", "C3:C7")]
    [TestCase("Name", "B3:B7")]
    [TestCase("Qty", "C3:C7")]
    public void Column_accepts_column_letter_or_field_name(string column, string expectedAddress)
    {
        // "Qty" is also a valid column letter, but the column is outside of the table. Thus the table uses the field name.
        using var wb = new XLWorkbook();
        var dataRange = CreateDataRange(wb);

        var rangeColumn = dataRange.Column(column);

        Assert.AreEqual(expectedAddress, rangeColumn.RangeAddress.ToString());
    }

    /// <summary>
    /// Create a table at B2:C7. The data range is B3:C7. The data rows 2 and 5 are empty.
    /// </summary>
    private static IXLTableRange CreateDataRange(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet();
        ws.Cell("B2").Value = "Name";
        ws.Cell("C2").Value = "Qty";
        ws.Cell("B3").Value = "Cake";
        ws.Cell("C3").Value = 7;
        ws.Cell("B5").Value = "Waffle";
        ws.Cell("C5").Value = 4;
        ws.Cell("B6").Value = "Pie";
        ws.Cell("C6").Value = 9;
        var table = ws.Range("B2:C7").CreateTable("Products");
        return table.DataRange;
    }

    private static string[] GetAddresses(IXLTableRows rows)
    {
        return rows.Select(r => r.RangeAddress.ToString()).ToArray();
    }

    [Test]
    public void First_and_last_row_used_are_null_for_table_range_without_used_rows()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "First";
        var table = ws.Range("A1:A3").CreateTable();
        var dataRange = table.DataRange;

        Assert.IsNull(dataRange.FirstRowUsed());
        Assert.IsNull(dataRange.LastRowUsed());
        Assert.IsNull(dataRange.FirstRowUsed(XLCellsUsedOptions.All));
        Assert.IsNull(dataRange.LastRowUsed(XLCellsUsedOptions.All));
        Assert.AreEqual(2, dataRange.FirstRow().RowNumber());
        Assert.AreEqual(3, dataRange.LastRow().RowNumber());
        Assert.IsNull(dataRange.FirstRow(r => r.RowNumber() > 3));
        Assert.IsNull(dataRange.LastRow(r => r.RowNumber() < 2));
    }
}
