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
    /// Create a file with a data bar that has the id in uppercase and the <paramref name="gradient"/> attribute of the extension.
    /// </summary>
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
