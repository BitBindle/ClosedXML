using System.IO;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ClosedXML.Tests.Excel.ConditionalFormats;

internal class ConditionalFormatDataBarTests
{
    [Test]
    [Issue("ClosedXML/ClosedXML#386")]
    [Issue("ClosedXML/ClosedXML#1686")]
    public void Gradient_of_data_bar_is_saved_and_loaded()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            var defaultBar = ws.Range("A1:A3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
            var solidBar = ws.Range("B1:B3").AddConditionalFormat().DataBar(XLColor.Red).Minimum(XLCFContentType.Number, 0).Maximum(XLCFContentType.Number, 1).SetGradient(false);
            var gradientBar = ws.Range("C1:C3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().Maximum(XLCFContentType.Percent, 90.0).SetGradient();

            Assert.IsTrue(defaultBar.Gradient);
            Assert.IsFalse(solidBar.Gradient);
            Assert.IsTrue(gradientBar.Gradient);
            wb.SaveAs(ms);
        }

        CollectionAssert.AreEqual(new bool?[] { true, false, true }, ReadGradientAttributes(ms));

        using var loaded = new XLWorkbook(ms);
        var formats = loaded.Worksheets.Single().ConditionalFormats.OrderBy(cf => cf.Range.FirstCell().Address.ColumnNumber);
        CollectionAssert.AreEqual(new[] { true, false, true }, formats.Select(cf => cf.Gradient));
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#386")]
    [Issue("ClosedXML/ClosedXML#1686")]
    public void Solid_data_bar_from_file_stays_solid_on_save()
    {
        // Excel writes the gradient attribute only for a solid fill.
        using var input = CreateFileWithDataBar(gradient: false);
        using var output = new MemoryStream();

        using (var wb = new XLWorkbook(input))
        {
            Assert.IsFalse(wb.Worksheets.Single().ConditionalFormats.Single().Gradient);
            wb.SaveAs(output);
        }

        CollectionAssert.AreEqual(new bool?[] { false }, ReadGradientAttributes(output));
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#386")]
    [Issue("ClosedXML/ClosedXML#1686")]
    public void Fill_of_data_bar_from_file_can_change()
    {
        // The fluent SetGradient is available only when a new data bar is made.
        using var input = CreateFileWithDataBar(gradient: null);
        using var output = new MemoryStream();

        using (var wb = new XLWorkbook(input))
        {
            wb.Worksheets.Single().ConditionalFormats.Single().Gradient = false;
            wb.SaveAs(output);
        }

        CollectionAssert.AreEqual(new bool?[] { false }, ReadGradientAttributes(output));
    }

    [Test]
    [Issue("ClosedXML/ClosedXML#386")]
    [Issue("ClosedXML/ClosedXML#1686")]
    public void Load_reads_extension_of_data_bar_with_uppercase_id()
    {
        // Excel writes the id of a rule in uppercase. ClosedXML writes it in lowercase.
        using var ms = CreateFileWithDataBar(gradient: null);

        using var wb = new XLWorkbook(ms);
        var cf = wb.Worksheets.Single().ConditionalFormats.Single();

        Assert.AreEqual(XLColor.Blue, cf.Colors[2]);
        Assert.IsTrue(cf.Gradient);
    }

    [Test]
    public void Load_accepts_data_bar_without_negative_fill_color()
    {
        // The schema makes the negative fill color optional.
        using var ms = CreateFileWithDataBar(gradient: false, removeNegativeFillColor: true);

        using var wb = new XLWorkbook(ms);
        var cf = wb.Worksheets.Single().ConditionalFormats.Single();

        Assert.AreEqual(1, cf.Colors.Count);
        Assert.AreEqual(XLColor.Red, cf.Colors[1]);
        Assert.IsFalse(cf.Gradient);
    }

    [Test]
    public void Copy_keeps_gradient_of_data_bar()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue().SetGradient(false);

        var copy = ws.CopyTo("Copy");

        Assert.IsFalse(copy.ConditionalFormats.Single().Gradient);
    }

    [Test]
    public void Data_bars_with_different_fills_are_not_equal()
    {
        // The compare of formats without the range must not join a solid and a gradient data bar.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("B1:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("C1:C3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue().SetGradient(false);
        var formats = ws.ConditionalFormats.Cast<XLConditionalFormat>().ToArray();
        var comparer = XLConditionalFormat.NoRangeComparer;

        Assert.IsTrue(comparer.Equals(formats[0], formats[1]));
        Assert.AreEqual(comparer.GetHashCode(formats[0]), comparer.GetHashCode(formats[1]));
        Assert.IsFalse(comparer.Equals(formats[0], formats[2]));
    }

    [Test]
    [Issue("61")]
    public void Save_keeps_data_bar_options_from_file()
    {
        using var input = CreateFileWithDataBarOptions(X14.DataBarDirectionValues.RightToLeft, X14.DataBarAxisPositionValues.Middle);
        using var output = new MemoryStream();

        using (var wb = new XLWorkbook(input))
            wb.SaveAs(output);

        using var package = SpreadsheetDocument.Open(output, false);
        var dataBar = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<X14.DataBar>().Single();
        Assert.AreEqual(5u, dataBar.MinLength!.Value);
        Assert.AreEqual(95u, dataBar.MaxLength!.Value);
        Assert.IsTrue(dataBar.Border!.Value);
        Assert.IsTrue(dataBar.NegativeBarColorSameAsPositive!.Value);
        Assert.IsFalse(dataBar.NegativeBarBorderColorSameAsPositive!.Value);
        Assert.AreEqual(X14.DataBarDirectionValues.RightToLeft, dataBar.Direction!.Value);
        Assert.AreEqual(X14.DataBarAxisPositionValues.Middle, dataBar.AxisPosition!.Value);

        // The schema gives the order of the child elements.
        CollectionAssert.AreEqual(
            new[] { "cfvo", "cfvo", "borderColor", "negativeFillColor", "negativeBorderColor", "axisColor" },
            dataBar.ChildElements.Select(e => e.LocalName));
        Assert.AreEqual("FF638EC6", dataBar.GetFirstChild<X14.BorderColor>()!.Rgb!.Value);
        var negativeBorderColor = dataBar.GetFirstChild<X14.NegativeBorderColor>()!;
        Assert.AreEqual(5u, negativeBorderColor.Theme!.Value);
        Assert.AreEqual(-0.25, negativeBorderColor.Tint!.Value);
        Assert.AreEqual("FF00B050", dataBar.GetFirstChild<X14.BarAxisColor>()!.Rgb!.Value);
    }

    [TestCase("leftToRight", "none")]
    [TestCase("rightToLeft", "middle")]
    [TestCase("context", "automatic")]
    [Issue("61")]
    public void Save_keeps_direction_and_axis_position_of_data_bar(string direction, string axisPosition)
    {
        using var input = CreateFileWithDataBarOptions(new X14.DataBarDirectionValues(direction), new X14.DataBarAxisPositionValues(axisPosition));
        using var output = new MemoryStream();

        using (var wb = new XLWorkbook(input))
            wb.SaveAs(output);

        // The save does not write an attribute with the schema default value.
        using var package = SpreadsheetDocument.Open(output, false);
        var dataBar = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<X14.DataBar>().Single();
        Assert.AreEqual(direction == "context" ? null : direction, dataBar.Direction?.InnerText);
        Assert.AreEqual(axisPosition == "automatic" ? null : axisPosition, dataBar.AxisPosition?.InnerText);
    }

    [Test]
    [Issue("61")]
    public void Load_uses_schema_defaults_for_missing_data_bar_options()
    {
        // The schema default of minLength is 10 and of maxLength is 90. The axis color is optional.
        using var input = CreateFileWithDataBar(gradient: null);
        using (var package = SpreadsheetDocument.Open(input, true))
        {
            var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            var dataBar = worksheet.Descendants<X14.DataBar>().Single();
            dataBar.MinLength = null;
            dataBar.MaxLength = null;
            dataBar.GetFirstChild<X14.BarAxisColor>()!.Remove();
            worksheet.Save();
        }

        input.Position = 0;
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(input))
            wb.SaveAs(output);

        using var savedPackage = SpreadsheetDocument.Open(output, false);
        var savedDataBar = savedPackage.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<X14.DataBar>().Single();
        Assert.AreEqual(10u, savedDataBar.MinLength!.Value);
        Assert.AreEqual(90u, savedDataBar.MaxLength!.Value);
        Assert.IsNull(savedDataBar.GetFirstChild<X14.BarAxisColor>());
        Assert.IsNull(savedDataBar.Border);
        Assert.IsNull(savedDataBar.NegativeBarColorSameAsPositive);
        Assert.IsNull(savedDataBar.NegativeBarBorderColorSameAsPositive);
    }

    [Test]
    [Issue("61")]
    public void Copy_keeps_data_bar_options_from_file()
    {
        using var input = CreateFileWithDataBarOptions(X14.DataBarDirectionValues.LeftToRight, X14.DataBarAxisPositionValues.None);
        using var wb = new XLWorkbook(input);
        var ws = wb.Worksheets.Single();

        var copy = ws.CopyTo("Copy");

        var original = (XLConditionalFormat)ws.ConditionalFormats.Single();
        var copied = (XLConditionalFormat)copy.ConditionalFormats.Single();
        Assert.AreEqual(original.DataBarOptions, copied.DataBarOptions);
        Assert.IsTrue(copied.DataBarOptions.Border);
    }

    [Test]
    [Issue("61")]
    public void Data_bars_with_different_options_are_not_equal()
    {
        // The compare of formats without the range must not join data bars that look different.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:A3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        ws.Range("B1:B3").AddConditionalFormat().DataBar(XLColor.Red).LowestValue().HighestValue();
        var formats = ws.ConditionalFormats.Cast<XLConditionalFormat>().ToArray();
        formats[1].DataBarOptions = formats[1].DataBarOptions with { Border = true };

        Assert.IsFalse(XLConditionalFormat.NoRangeComparer.Equals(formats[0], formats[1]));
    }

    [Test]
    [Issue("69")]
    public void Save_keeps_automatic_axis_color_of_data_bar()
    {
        using var input = CreateFileWithDataBar(gradient: null);
        using (var package = SpreadsheetDocument.Open(input, true))
        {
            var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            var axisColor = worksheet.Descendants<X14.BarAxisColor>().Single();
            axisColor.Rgb = null;
            axisColor.Auto = true;
            worksheet.Save();
        }

        input.Position = 0;
        using var output = new MemoryStream();
        using (var wb = new XLWorkbook(input))
        {
            var options = ((XLConditionalFormat)wb.Worksheets.Single().ConditionalFormats.Single()).DataBarOptions;
            Assert.AreEqual(XLColor.Automatic, options.AxisColor);
            wb.SaveAs(output);
        }

        using var savedPackage = SpreadsheetDocument.Open(output, false);
        var savedAxisColor = savedPackage.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<X14.BarAxisColor>().Single();
        Assert.IsTrue(savedAxisColor.Auto!.Value);
        Assert.IsNull(savedAxisColor.Rgb);
    }

    private static bool?[] ReadGradientAttributes(MemoryStream ms)
    {
        using (var package = SpreadsheetDocument.Open(ms, false))
        {
            var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            var result = worksheet.Descendants<X14.DataBar>().Select(dataBar => dataBar.Gradient?.Value).ToArray();
            ms.Position = 0;
            return result;
        }
    }

    /// <summary>
    /// Create a file with a data bar that has the options of Excel 2010 that the public API does not set.
    /// </summary>
    private static MemoryStream CreateFileWithDataBarOptions(X14.DataBarDirectionValues direction, X14.DataBarAxisPositionValues axisPosition)
    {
        var ms = CreateFileWithDataBar(gradient: null);
        using (var package = SpreadsheetDocument.Open(ms, true))
        {
            var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            var dataBar = worksheet.Descendants<X14.DataBar>().Single();
            dataBar.MinLength = 5;
            dataBar.MaxLength = 95;
            dataBar.Border = true;
            dataBar.NegativeBarColorSameAsPositive = true;
            dataBar.NegativeBarBorderColorSameAsPositive = false;
            dataBar.Direction = direction;
            dataBar.AxisPosition = axisPosition;

            var negativeFillColor = dataBar.GetFirstChild<X14.NegativeFillColor>()!;
            dataBar.InsertBefore(new X14.BorderColor { Rgb = "FF638EC6" }, negativeFillColor);
            dataBar.InsertAfter(new X14.NegativeBorderColor { Theme = 5, Tint = -0.25 }, negativeFillColor);
            dataBar.GetFirstChild<X14.BarAxisColor>()!.Rgb = "FF00B050";
            worksheet.Save();
        }

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Create a file with a data bar that has the id in uppercase and the <paramref name="gradient"/> attribute of the extension.
    /// </summary>
    /// <param name="gradient">The gradient attribute of the extension. A null value removes the attribute.</param>
    /// <param name="removeNegativeFillColor">If <c>true</c>, the extension has no negative fill color.</param>
    private static MemoryStream CreateFileWithDataBar(bool? gradient, bool removeNegativeFillColor = false)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Range("A1:A3").AddConditionalFormat().DataBar(XLColor.Red, XLColor.Blue).LowestValue().HighestValue();
            wb.SaveAs(ms);
        }

        using (var package = SpreadsheetDocument.Open(ms, true))
        {
            var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            foreach (var id in worksheet.Descendants<X14.Id>())
                id.Text = id.Text.ToUpperInvariant();
            foreach (var rule in worksheet.Descendants<X14.ConditionalFormattingRule>())
                rule.Id = rule.Id!.Value!.ToUpperInvariant();
            foreach (var dataBar in worksheet.Descendants<X14.DataBar>())
                dataBar.Gradient = gradient;
            if (removeNegativeFillColor)
            {
                foreach (var color in worksheet.Descendants<X14.NegativeFillColor>().ToList())
                    color.Remove();
            }
            worksheet.Save();
        }

        ms.Position = 0;
        return ms;
    }
}
