using System;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Charts;

/// <summary>
/// Test of <see cref="XLChart"/> and its base class <see cref="XLDrawing{T}"/>.
/// </summary>
[TestFixture]
internal class XLChartTests
{
    [Test]
    public void New_chart_has_default_values()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();

        var chart = new XLChart(ws);

        Assert.That(chart.ZOrder, Is.EqualTo(1));
        Assert.That(chart.ShapeId, Is.GreaterThan(0));
        Assert.That(chart.RightAngleAxes, Is.True);
        Assert.That(chart.Visible, Is.False);
        Assert.That(chart.Position, Is.Not.Null);
        Assert.That(chart.Style, Is.Not.Null);
    }

    [Test]
    public void New_chart_is_above_existing_charts_of_worksheet()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        var first = new XLChart(ws);
        first.SetZOrder(5);
        ws.Charts.Add(first);

        var second = new XLChart(ws);

        Assert.That(second.ZOrder, Is.EqualTo(6));
        Assert.That(second.ShapeId, Is.Not.EqualTo(first.ShapeId));
        Assert.That(ws.Charts.Single(), Is.SameAs(first));
    }

    [Test]
    public void Chart_properties_can_be_set()
    {
        using var wb = new XLWorkbook();
        var chart = new XLChart((XLWorksheet)wb.AddWorksheet());

        IXLChart[] results =
        [
            chart.SetRightAngleAxes(false),
            chart.SetChartType(XLChartType.Pie),
            chart.SetVisible(true),
            chart.SetName("Sales"),
            chart.SetDescription("Sales of year"),
            chart.SetZOrder(3),
            chart.SetHorizontalFlip(true),
            chart.SetVerticalFlip(true),
            chart.SetRotation(45),
            chart.SetOffsetX(10),
            chart.SetOffsetY(20),
            chart.SetExtentLength(300),
            chart.SetExtentWidth(400),
        ];

        Assert.That(results, Is.All.SameAs(chart));
        Assert.That(chart.RightAngleAxes, Is.False);
        Assert.That(chart.ChartType, Is.EqualTo(XLChartType.Pie));
        Assert.That(chart.Visible, Is.True);
        Assert.That(chart.Name, Is.EqualTo("Sales"));
        Assert.That(chart.Description, Is.EqualTo("Sales of year"));
        Assert.That(chart.ZOrder, Is.EqualTo(3));
        Assert.That(chart.HorizontalFlip, Is.True);
        Assert.That(chart.VerticalFlip, Is.True);
        Assert.That(chart.Rotation, Is.EqualTo(45));
        Assert.That(chart.OffsetX, Is.EqualTo(10));
        Assert.That(chart.OffsetY, Is.EqualTo(20));
        Assert.That(chart.ExtentLength, Is.EqualTo(300));
        Assert.That(chart.ExtentWidth, Is.EqualTo(400));
    }

    [Test]
    public void Chart_flags_can_be_set_without_value()
    {
        using var wb = new XLWorkbook();
        var chart = new XLChart((XLWorksheet)wb.AddWorksheet());
        chart.SetRightAngleAxes(false);

        chart.SetRightAngleAxes();
        chart.SetVisible();
        chart.SetHorizontalFlip();
        chart.SetVerticalFlip();

        Assert.That(chart.RightAngleAxes, Is.True);
        Assert.That(chart.Visible, Is.True);
        Assert.That(chart.HorizontalFlip, Is.True);
        Assert.That(chart.VerticalFlip, Is.True);
    }

    [TestCase(XLChartType.BarClustered3D)]
    [TestCase(XLChartType.ColumnStacked3D)]
    public void Category_of_3D_bar_chart_is_3D_bar(XLChartType chartType)
    {
        using var wb = new XLWorkbook();
        var chart = new XLChart((XLWorksheet)wb.AddWorksheet()) { ChartType = chartType };

        Assert.That(chart.ChartTypeCategory, Is.EqualTo(XLChartTypeCategory.Bar3D));
    }

    [Test]
    public void Category_of_other_chart_types_is_not_implemented()
    {
        // This test records a missing feature, not a requirement. If a category is added, change this test.
        using var wb = new XLWorkbook();
        var chart = new XLChart((XLWorksheet)wb.AddWorksheet()) { ChartType = XLChartType.Pie };

        Assert.Throws<NotImplementedException>(() => _ = chart.ChartTypeCategory);
    }

    [TestCase(XLChartType.BarClustered, XLBarOrientation.Horizontal)]
    [TestCase(XLChartType.PyramidHorizontalStacked, XLBarOrientation.Horizontal)]
    [TestCase(XLChartType.ColumnClustered, XLBarOrientation.Vertical)]
    public void Bar_orientation_depends_on_chart_type(XLChartType chartType, XLBarOrientation expected)
    {
        using var wb = new XLWorkbook();
        var chart = new XLChart((XLWorksheet)wb.AddWorksheet()) { ChartType = chartType };

        Assert.That(chart.BarOrientation, Is.EqualTo(expected));
    }

    [TestCase(XLChartType.ColumnClustered, XLBarGrouping.Clustered)]
    [TestCase(XLChartType.BarStacked100Percent, XLBarGrouping.Percent)]
    [TestCase(XLChartType.AreaStacked, XLBarGrouping.Stacked)]
    [TestCase(XLChartType.Column3D, XLBarGrouping.Standard)]
    public void Bar_grouping_depends_on_chart_type(XLChartType chartType, XLBarGrouping expected)
    {
        using var wb = new XLWorkbook();
        var chart = new XLChart((XLWorksheet)wb.AddWorksheet()) { ChartType = chartType };

        Assert.That(chart.BarGrouping, Is.EqualTo(expected));
    }
}
