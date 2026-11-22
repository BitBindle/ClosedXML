using System;
using System.Linq;
using ClosedXML.Excel;
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
