using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.PivotTables;

/// <summary>
/// Tests methods of interface <see cref="IXLPivotField"/> implemented through <see cref="XLPivotTablePageField"/>.
/// </summary>
[TestFixture]
internal class XLPivotTablePageFieldTests
{
    [Test]
    public void Setters_change_pivot_field_of_filter()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        filter
            .SetCustomName("Town")
            .SetSubtotalCaption("Town total")
            .SetIncludeNewItemsInFilter(true)
            .SetSubtotalsAtTop(false)
            .SetRepeatItemLabels(true)
            .SetInsertBlankLines(true)
            .SetShowBlankItems(false)
            .SetInsertPageBreaks(true)
            .SetCollapsed(true)
            .SetSort(XLPivotSortType.Descending);

        var pivotField = ((XLPivotTable)pt).PivotFields[1];
        Assert.AreEqual("Town", pivotField.Name);
        Assert.AreEqual("Town", filter.CustomName);
        Assert.AreEqual("Town total", filter.SubtotalCaption);
        Assert.True(filter.IncludeNewItemsInFilter);
        Assert.False(filter.SubtotalsAtTop);
        Assert.True(filter.RepeatItemLabels);
        Assert.True(filter.InsertBlankLines);
        Assert.False(filter.ShowBlankItems);
        Assert.True(filter.InsertPageBreaks);
        Assert.True(filter.Collapsed);
        Assert.AreEqual(XLPivotSortType.Descending, filter.SortType);
    }

    [Test]
    public void SubtotalsAtTop_without_value_puts_subtotals_at_top()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");
        filter.SubtotalsAtTop = false;

        filter.SubtotalsAtTop = null;

        Assert.True(filter.SubtotalsAtTop);
    }

    [Test]
    public void AddSubtotal_adds_function_to_subtotals()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        filter.AddSubtotal(XLSubtotalFunction.Sum);

        CollectionAssert.AreEquivalent(new[] { XLSubtotalFunction.Automatic, XLSubtotalFunction.Sum }, filter.Subtotals);
    }

    [TestCase(XLPivotLayout.Compact, true, true)]
    [TestCase(XLPivotLayout.Outline, true, false)]
    [TestCase(XLPivotLayout.Tabular, false, false)]
    public void SetLayout_sets_outline_and_compact_flags(XLPivotLayout layout, bool outline, bool compact)
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        filter.SetLayout(layout);

        Assert.AreEqual(outline, filter.Outline);
        Assert.AreEqual(compact, filter.Compact);
    }

    [Test]
    public void SelectedValues_without_selection_contains_all_values()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        Assert.AreEqual(new XLCellValue[] { "Tokyo", "Paris", "Rome", "Oslo" }, filter.SelectedValues);
    }

    [Test]
    public void AddSelectedValues_selects_only_added_values()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        filter.AddSelectedValues(new XLCellValue[] { "Tokyo", "Rome", "Oslo" });

        Assert.AreEqual(new XLCellValue[] { "Tokyo", "Rome", "Oslo" }, filter.SelectedValues);
    }

    [Test]
    [Issue("31")]
    public void SelectedValues_with_one_selected_value_contains_only_that_value()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        filter.AddSelectedValue("Rome");

        Assert.AreEqual(new XLCellValue[] { "Rome" }, filter.SelectedValues);
    }

    [Test]
    [Issue("31")]
    public void AddSelectedValue_after_one_selected_value_selects_both_values()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");
        filter.AddSelectedValue("Rome");

        filter.AddSelectedValue("Tokyo");

        Assert.AreEqual(new XLCellValue[] { "Tokyo", "Rome" }, filter.SelectedValues);
    }

    [Test]
    [Issue("31")]
    public void One_selected_value_is_loaded_from_position_in_pivot_field_items()
    {
        // ECMA-376 defines pageField@item as an index into pivotField.items, not into the shared items.
        // The field items are in reverse order, thus the item at position 0 has the value of shared item 3.
        using var ms = CreateFileWithReversedCityItems(selectedItem: 0);

        using var wb = new XLWorkbook(ms);
        var filter = wb.Worksheet(1).PivotTable("pt").ReportFilters.Single();

        Assert.AreEqual(new XLCellValue[] { "Oslo" }, filter.SelectedValues);
    }

    [Test]
    [Issue("31")]
    public void One_selected_value_is_saved_as_position_in_pivot_field_items()
    {
        // The field items are in reverse order, thus shared item 2 "Rome" is at position 1.
        using var ms = CreateFileWithReversedCityItems(selectedItem: null);
        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet(1).PivotTable("pt").ReportFilters.Single().AddSelectedValue("Rome");
            wb.Save();
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var definition = document.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            var pageField = definition.PageFields!.Elements<PageField>().Single();
            Assert.AreEqual(1, pageField.Item!.Value);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var filter = wb.Worksheet(1).PivotTable("pt").ReportFilters.Single();
            Assert.AreEqual(new XLCellValue[] { "Rome" }, filter.SelectedValues);
        }
    }

    [Test]
    [Issue("31")]
    public void AddSelectedValue_after_one_loaded_selected_value_keeps_loaded_value()
    {
        using var ms = CreateFileWithReversedCityItems(selectedItem: 0);
        using var wb = new XLWorkbook(ms);
        var filter = wb.Worksheet(1).PivotTable("pt").ReportFilters.Single();

        filter.AddSelectedValue("Tokyo");

        CollectionAssert.AreEquivalent(new XLCellValue[] { "Oslo", "Tokyo" }, filter.SelectedValues);
    }

    [Test]
    [Issue("41")]
    public void AddSelectedValue_adds_missing_field_item_before_subtotal_items()
    {
        // Excel tries to repair a workbook if a subtotal item is not at the end of the field items.
        using var ms = CreateFileWithoutRomeItem();
        using (var wb = new XLWorkbook(ms))
        {
            wb.Worksheet(1).PivotTable("pt").ReportFilters.Single().AddSelectedValue("Rome");
            wb.Save();
        }

        using (var document = SpreadsheetDocument.Open(ms, false))
        {
            var definition = document.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            var cityItems = definition.PivotFields!.Elements<PivotField>().ElementAt(1).Items!.Elements<Item>().ToList();
            Assert.AreEqual(new uint?[] { 0, 1, 3, 2, null }, cityItems.Select(i => i.Index?.Value));
            Assert.AreEqual(ItemValues.Default, cityItems.Last().ItemType!.Value);
            var pageField = definition.PageFields!.Elements<PageField>().Single();
            Assert.AreEqual(3, pageField.Item!.Value);
        }

        using (var wb = new XLWorkbook(ms))
        {
            var filter = wb.Worksheet(1).PivotTable("pt").ReportFilters.Single();
            Assert.AreEqual(new XLCellValue[] { "Rome" }, filter.SelectedValues);
        }
    }

    [Test]
    [Issue("41")]
    public void AddSelectedValue_with_multiple_selection_adds_missing_field_item_before_subtotal_items()
    {
        using var ms = CreateFileWithoutRomeItem();
        using var wb = new XLWorkbook(ms);
        var pt = (XLPivotTable)wb.Worksheet(1).PivotTable("pt");
        var filter = pt.ReportFilters.Single();

        filter.AddSelectedValue("Paris");
        filter.AddSelectedValue("Rome");

        var cityItems = pt.PivotFields[1].Items;
        Assert.AreEqual(new int?[] { 0, 1, 3, 2, null }, cityItems.Select(i => i.ItemIndex));
        Assert.AreEqual(XLPivotItemType.Default, cityItems.Last().ItemType);
        CollectionAssert.AreEquivalent(new XLCellValue[] { "Paris", "Rome" }, filter.SelectedValues);
    }

    [Test]
    [Issue("41")]
    public void GetAllItems_adds_missing_field_item_before_subtotal_items()
    {
        using var ms = CreateFileWithoutRomeItem();
        using var wb = new XLWorkbook(ms);
        var pt = (XLPivotTable)wb.Worksheet(1).PivotTable("pt");
        var cityField = pt.PivotFields[1];

        var items = cityField.GetAllItems(value => value.GetText().Contains('o'));

        Assert.AreEqual(new int?[] { 0, 2, 3 }, items.Select(i => i.ItemIndex));
        Assert.AreEqual(new int?[] { 0, 1, 3, 2, null }, cityField.Items.Select(i => i.ItemIndex));
        Assert.AreEqual(XLPivotItemType.Default, cityField.Items.Last().ItemType);
    }

    [Test]
    public void Filter_field_is_only_in_filter_list()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        Assert.AreEqual("City", filter.SourceName);
        Assert.False(filter.IsOnRowAxis);
        Assert.False(filter.IsOnColumnAxis);
        Assert.True(filter.IsInFilterList);
        Assert.AreEqual(1, filter.Offset);
    }

    [Test]
    public void StyleFormats_of_filter_field_are_not_implemented()
    {
        using var wb = new XLWorkbook();
        var pt = CreatePivotTable(wb);
        var filter = pt.ReportFilters.Add("City");

        Assert.Throws<NotImplementedException>(() => _ = filter.StyleFormats);
    }

    /// <summary>
    /// Creates a file with a City filter. The data items of the City field are in reverse order
    /// of the shared items. Excel can save such an order, for example after a sort of the items.
    /// </summary>
    private static MemoryStream CreateFileWithReversedCityItems(uint? selectedItem)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.ReportFilters.Add("City");
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

            if (selectedItem is not null)
                definition.PageFields!.Elements<PageField>().Single().Item = selectedItem.Value;
        }

        return ms;
    }

    /// <summary>
    /// Creates a file with a City filter. The City field has no item for the shared item "Rome".
    /// Excel does not save such a field, but other producers can.
    /// </summary>
    private static MemoryStream CreateFileWithoutRomeItem()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var pt = CreatePivotTable(wb);
            pt.ReportFilters.Add("City");
            wb.SaveAs(ms);
        }

        using (var document = SpreadsheetDocument.Open(ms, true))
        {
            var definition = document.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single().PivotTableDefinition!;
            var cityItems = definition.PivotFields!.Elements<PivotField>().ElementAt(1).Items!;
            cityItems.Elements<Item>().Single(i => i.Index?.Value == 2).Remove();
            cityItems.Count = (uint)cityItems.Elements<Item>().Count();
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
        return ws.PivotTables.Add("pt", ws.Cell("F2"), data);
    }
}
