using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.CalcEngine;

/// <summary>
/// Tests of the internal methods of <see cref="Reference"/>.
/// </summary>
[TestFixture]
internal class ReferenceTests
{
    [Test]
    public void Area_must_be_normalized()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var area = new XLRangeAddress(new XLAddress(ws, 3, 3, false, false), new XLAddress(ws, 1, 1, false, false));

        Assert.Throws<ArgumentException>(() => _ = new Reference(area));
    }

    [Test]
    public void Reference_must_have_at_least_one_area()
    {
        using var wb = new XLWorkbook();

        Assert.Throws<ArgumentException>(() => _ = new Reference(new List<XLRangeAddress>()));
        Assert.Throws<ArgumentException>(() => _ = new Reference(new XLRanges(wb)));
    }

    [Test]
    public void Reference_from_ranges_has_area_of_each_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var reference = new Reference(ws.Ranges("A1:B2,D4"));

        Assert.AreEqual(new[] { "A1:B2", "D4:D4" }, reference.Areas.Select(a => a.ToString()).ToArray());
        Assert.AreEqual(5, reference.NumberOfCells);
    }

    [Test]
    public void Intersection_of_areas_on_same_sheet_contains_shared_cells()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var ctx = CreateContext(wb, ws);

        var result = Reference.Intersect(Create(ws, "A1:C3", "E1:F4"), Create(ws, "B2:E3"), ctx);

        Assert.True(result.TryPickT0(out var intersection, out _));
        Assert.AreEqual(new[] { "B2:C3", "E2:E3" }, intersection!.Areas.Select(a => a.ToString()).ToArray());
    }

    [Test]
    [Issue("48")]
    public void Intersection_intersects_left_area_with_each_right_area()
    {
        // Excel intersects each left area with each right area. The result is the union of these intersections.
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var ctx = CreateContext(wb, ws);

        var result = Reference.Intersect(Create(ws, "A1:B2"), Create(ws, "A1", "B2"), ctx);

        Assert.True(result.TryPickT0(out var intersection, out _));
        Assert.AreEqual(new[] { "A1:A1", "B2:B2" }, intersection!.Areas.Select(a => a.ToString()).ToArray());
    }

    [Test]
    public void Intersection_without_shared_cells_is_null_error()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var ctx = CreateContext(wb, ws);

        var result = Reference.Intersect(Create(ws, "A1:B2"), Create(ws, "C3:D4"), ctx);

        Assert.False(result.TryPickT0(out _, out var error));
        Assert.AreEqual(XLError.NullValue, error);
    }

    [Test]
    public void Intersection_of_areas_on_different_sheets_is_value_error()
    {
        using var wb = new XLWorkbook();
        var ws1 = (XLWorksheet)wb.AddWorksheet();
        var ws2 = (XLWorksheet)wb.AddWorksheet();
        var ctx = CreateContext(wb, ws1);

        var result = Reference.Intersect(Create(ws1, "A1:B2"), Create(ws2, "A1:B2"), ctx);

        Assert.False(result.TryPickT0(out _, out var error));
        Assert.AreEqual(XLError.IncompatibleValue, error);
    }

    [Test]
    public void Single_area_can_be_converted_to_array_of_cell_values()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("B2").Value = 1;
        ws.Cell("C3").Value = "text";
        var ctx = CreateContext(wb, ws);

        var result = Create(ws, "B2:C3").ToArray(ctx);

        Assert.True(result.TryPickT0(out var array, out _));
        Assert.AreEqual(2, array!.Width);
        Assert.AreEqual(2, array.Height);
        Assert.True(array[0, 0].TryPickNumber(out var number));
        Assert.AreEqual(1, number);
        Assert.True(array[0, 1].IsBlank);
        Assert.True(array[1, 1].IsText);
    }

    [Test]
    public void Apply_gives_array_with_result_for_each_cell()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("B2").Value = 4;
        var ctx = CreateContext(wb, ws);

        var result = Create(ws, "A1:B2").Apply(v => v.TryPickNumber(out var n) ? n * 10 : XLError.IncompatibleValue, ctx);

        Assert.True(result.TryPickT0(out var array, out _));
        var values = new double[2, 2];
        for (var y = 0; y < 2; y++)
        {
            for (var x = 0; x < 2; x++)
            {
                Assert.True(array![y, x].TryPickNumber(out var number));
                values[y, x] = number;
            }
        }

        Assert.AreEqual(new double[,] { { 10, 20 }, { 30, 40 } }, values);
    }

    [Test]
    public void Array_operations_need_single_area()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var ctx = CreateContext(wb, ws);
        var reference = Create(ws, "A1:B2", "D4");

        Assert.False(reference.ToArray(ctx).TryPickT0(out _, out var toArrayError));
        Assert.AreEqual(XLError.IncompatibleValue, toArrayError);
        Assert.False(reference.Apply(v => v, ctx).TryPickT0(out _, out var applyError));
        Assert.AreEqual(XLError.IncompatibleValue, applyError);
    }

    [Test]
    public void Implicit_intersection_of_single_cell_is_same_reference()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var reference = Create(ws, "C3");

        var result = reference.ImplicitIntersection(ws.Cell("A1").Address);

        Assert.True(result.TryPickT0(out var intersection, out _));
        Assert.AreSame(reference, intersection);
    }

    [Test]
    public void Single_cell_value_needs_single_cell_reference()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        ws.Cell("B2").Value = 5;
        var ctx = CreateContext(wb, ws);

        Assert.True(Create(ws, "B2").TryGetSingleCellValue(out var value, ctx));
        Assert.True(value.TryPickNumber(out var number));
        Assert.AreEqual(5, number);
        Assert.False(Create(ws, "B2:B3").TryGetSingleCellValue(out _, ctx));
    }

    private static Reference Create(XLWorksheet ws, params string[] areas)
    {
        return new Reference(areas.Select(a => (XLRangeAddress)ws.Range(a).RangeAddress).ToList());
    }

    private static CalcContext CreateContext(XLWorkbook wb, XLWorksheet ws)
    {
        return new CalcContext(wb.CalcEngine, CultureInfo.InvariantCulture, wb, ws, null);
    }
}
