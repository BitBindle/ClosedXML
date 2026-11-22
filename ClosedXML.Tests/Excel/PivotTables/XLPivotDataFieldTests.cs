using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables;

/// <summary>
/// Tests methods of interface <see cref="IXLPivotValue"/> implemented through <see cref="XLPivotDataField"/>.
/// </summary>
[TestFixture]
internal class XLPivotDataFieldTests
{
    [Test]
    public void BaseItemValue_is_saved_as_position_in_items_of_base_field()
    {
        // ECMA-376 defines dataField@baseItem as an index into the items of the base field.
        // The City items are in reverse order, thus shared item 2 "Rome" is at position 1.
        using var ms = CreateFileWithReversedCityItems();
        using (var wb = new XLWorkbook(ms))
        {
            var value = wb.Worksheet(1).PivotTable("pt").Values.Single();
            value.ShowAsDifferenceFrom("City").And("Rome");

            Assert.AreEqual(XLCellValue.FromObject("Rome"), value.BaseItemValue);
            wb.Save();
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var definition = document.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            var dataField = definition.DataFields!.Elements<DataField>().Single();
            Assert.AreEqual(1, dataField.BaseItem!.Value);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var value = wb.Worksheet(1).PivotTable("pt").Values.Single();
            Assert.AreEqual(XLCellValue.FromObject("Rome"), value.BaseItemValue);
        }
    }

    [Test]
    public void BaseItemValue_needs_base_field()
    {
        using var wb = new XLWorkbook();
        var value = CreatePivotTable(wb).Values.Single();

        Assert.Throws<InvalidOperationException>(() => value.SetBaseItemValue("Rome"));
    }

    /// <summary>
    /// Creates a file with City on the row axis. The data items of the City field are in reverse
    /// order of the shared items. Excel can save such an order, for example after a sort of the items.
    /// </summary>
    private static MemoryStream CreateFileWithReversedCityItems()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            CreatePivotTable(wb);
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var definition = document.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            var cityItems = definition.PivotFields!.Elements<PivotField>().ElementAt(1).Items!;
            var dataItems = cityItems.Elements<Item>().Where(i => i.Index is not null).ToList();
            foreach (var item in dataItems)
                item.Remove();

            dataItems.Reverse();
            for (var i = 0; i < dataItems.Count; ++i)
                cityItems.InsertAt(dataItems[i], i);
        }

        return ms;
    }

    private static IXLPivotTable CreatePivotTable(XLWorkbook wb)
    {
        var ws = wb.AddWorksheet();
        var data = ws.Cell("A1").InsertData(new object[]
        {
            ("Name", "City", "Value"),
            ("Cake", "Tokyo", 7),
            ("Pie", "Paris", 4),
            ("Waffle", "Rome", 5),
            ("Bun", "Oslo", 9),
        });
        var pt = ws.PivotTables.Add("pt", ws.Cell("F2"), data);
        pt.RowLabels.Add("City");
        pt.Values.Add("Value");
        return pt;
    }
}
