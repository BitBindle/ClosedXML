using System;
using ClosedXML.Excel;
using ClosedXML.Excel.Formatting;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles;

/// <summary>
/// Test of <see cref="XLDxFormat"/>.
/// </summary>
[TestFixture]
internal class DxfFormatTests
{
    [Test]
    public void Assign_dxf_to_different_dxf()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var source = ws.Range("A1").AddConditionalFormat();
        source.Style.Fill.BackgroundColor = XLColor.Red;
        var target = ws.Range("B1").AddConditionalFormat();

        target.Style = source.Style;

        Assert.That(target.Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

        // Copy was deep, changes to the source don't affect the copy
        source.Style.Fill.BackgroundColor = XLColor.Green;
        Assert.That(target.Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
    }

    [Test]
    public void Cant_copy_cell_format_to_dxf()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        Assert.That(() => cf.Style = ws.Cell("B1").Style, Throws.TypeOf<NotSupportedException>());
    }

    [Test]
    public void IncludeQuotePrefix_always_returns_false()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();
        Assert.IsFalse(cf.Style.IncludeQuotePrefix);
    }

    [Test]
    public void IncludeQuotePrefix_cant_be_changed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();
        Assert.That(() => cf.Style.IncludeQuotePrefix = true, Throws.TypeOf<NotSupportedException>());
    }

    [TestCase(-1, "$0.00")]
    [TestCase(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_through_format(int numFmtId, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.NumberFormat.SetFormat(format);

        Assert.AreEqual(format, cf.Style.NumberFormat.Format);
        Assert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
    }

    [TestCase(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_through_number_format_id(int numFmtId, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.NumberFormat.SetNumberFormatId(numFmtId);

        Assert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
        Assert.AreEqual(format, cf.Style.NumberFormat.Format);
    }

    [TestCase(-1, "$0.00")]
    [TestCase(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_by_assigning_number_format(int numFmtId, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var numberFormat = ws.Range("A1").AddConditionalFormat().Style.NumberFormat.SetFormat(format).NumberFormat;
        var cf = ws.Range("A2").AddConditionalFormat();

        cf.Style.NumberFormat = numberFormat;

        Assert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
        Assert.AreEqual(format, cf.Style.NumberFormat.Format);
    }

    [Test]
    public void DateFormat_returns_number_format()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();
        cf.Style.NumberFormat.SetFormat("00.0");

        Assert.AreEqual("00.0", cf.Style.DateFormat.Format);
    }

    [Test]
    public void Protection_default_values_are_same_as_in_OOXML()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var protection = ws.Range("A1").AddConditionalFormat().Style.Protection;

        Assert.True(protection.Locked);
        Assert.False(protection.Hidden);
    }

    [Test]
    public void Protection_locked_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Protection.SetLocked(false);
        Assert.False(cf.Style.Protection.Locked);
    }

    [Test]
    public void Protection_hidden_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Protection.SetHidden(true);
        Assert.True(cf.Style.Protection.Hidden);
    }

    [Test]
    public void Alignment_default_values_are_same_as_in_OOXML()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var alignment = ws.Range("A1").AddConditionalFormat().Style.Alignment;

        Assert.AreEqual(XLAlignmentHorizontalValues.General, alignment.Horizontal);
        Assert.AreEqual(XLAlignmentVerticalValues.Bottom, alignment.Vertical);
        Assert.AreEqual(0, alignment.Indent);
        Assert.False(alignment.JustifyLastLine);
        Assert.AreEqual(XLAlignmentReadingOrderValues.ContextDependent, alignment.ReadingOrder);
        Assert.AreEqual(0, alignment.RelativeIndent);
        Assert.False(alignment.ShrinkToFit);
        Assert.AreEqual(0, alignment.TextRotation);
        Assert.False(alignment.WrapText);
        Assert.False(alignment.TopToBottom);
    }

    [Test]
    public void Alignment_values_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Range("A1").AddConditionalFormat().Style;

        var returnedStyle = style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(XLAlignmentVerticalValues.Top)
            .Alignment.SetIndent(2)
            .Alignment.SetJustifyLastLine()
            .Alignment.SetReadingOrder(XLAlignmentReadingOrderValues.RightToLeft)
            .Alignment.SetRelativeIndent(3)
            .Alignment.SetShrinkToFit()
            .Alignment.SetTextRotation(45)
            .Alignment.SetWrapText();

        Assert.AreSame(style, returnedStyle);
        var alignment = style.Alignment;
        Assert.AreEqual(XLAlignmentHorizontalValues.Center, alignment.Horizontal);
        Assert.AreEqual(XLAlignmentVerticalValues.Top, alignment.Vertical);
        Assert.AreEqual(2, alignment.Indent);
        Assert.True(alignment.JustifyLastLine);
        Assert.AreEqual(XLAlignmentReadingOrderValues.RightToLeft, alignment.ReadingOrder);
        Assert.AreEqual(3, alignment.RelativeIndent);
        Assert.True(alignment.ShrinkToFit);
        Assert.AreEqual(45, alignment.TextRotation);
        Assert.True(alignment.WrapText);
    }

    [Test]
    public void Alignment_boolean_values_can_be_cleared()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var alignment = ws.Range("A1").AddConditionalFormat().Style.Alignment;
        alignment.SetJustifyLastLine().Alignment.SetShrinkToFit().Alignment.SetWrapText();

        alignment.SetJustifyLastLine(false).Alignment.SetShrinkToFit(false).Alignment.SetWrapText(false);

        Assert.False(alignment.JustifyLastLine);
        Assert.False(alignment.ShrinkToFit);
        Assert.False(alignment.WrapText);
    }

    [Test]
    public void Alignment_top_to_bottom_is_vertical_text_rotation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var alignment = ws.Range("A1").AddConditionalFormat().Style.Alignment;

        alignment.SetTopToBottom();
        Assert.True(alignment.TopToBottom);
        Assert.AreEqual(TextRotation.VerticalText.Value, alignment.TextRotation);

        alignment.SetTopToBottom(false);
        Assert.False(alignment.TopToBottom);
        Assert.AreEqual(TextRotation.None.Value, alignment.TextRotation);
    }

    [Test]
    public void Alignment_can_be_set_by_assigning_cell_alignment()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Style.Alignment
            .SetHorizontal(XLAlignmentHorizontalValues.Right)
            .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
            .Alignment.SetIndent(4)
            .Alignment.SetReadingOrder(XLAlignmentReadingOrderValues.LeftToRight)
            .Alignment.SetShrinkToFit()
            .Alignment.SetTextRotation(30);
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Alignment = ws.Cell("B1").Style.Alignment;

        var alignment = cf.Style.Alignment;
        Assert.AreEqual(XLAlignmentHorizontalValues.Right, alignment.Horizontal);
        Assert.AreEqual(XLAlignmentVerticalValues.Center, alignment.Vertical);
        Assert.AreEqual(4, alignment.Indent);
        Assert.AreEqual(XLAlignmentReadingOrderValues.LeftToRight, alignment.ReadingOrder);
        Assert.True(alignment.ShrinkToFit);
        Assert.AreEqual(30, alignment.TextRotation);
        Assert.False(alignment.WrapText);
        Assert.False(alignment.JustifyLastLine);
    }

    [Test]
    public void Alignment_equality_is_not_supported()
    {
        // This test records a missing feature, not a requirement. If equality is added, change this test.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var alignment = ws.Range("A1").AddConditionalFormat().Style.Alignment;

        Assert.That(() => alignment.Equals(ws.Cell("A1").Style.Alignment), Throws.TypeOf<NotSupportedException>());
    }

    [Test]
    public void Border_default_values_are_same_as_in_OOXML()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var border = ws.Range("A1").AddConditionalFormat().Style.Border;

        Assert.AreEqual(XLBorderStyleValues.None, border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.None, border.RightBorder);
        Assert.AreEqual(XLBorderStyleValues.None, border.TopBorder);
        Assert.AreEqual(XLBorderStyleValues.None, border.BottomBorder);
        Assert.AreEqual(XLBorderStyleValues.None, border.DiagonalBorder);
        Assert.AreEqual(XLColor.Automatic, border.LeftBorderColor);
        Assert.AreEqual(XLColor.Automatic, border.RightBorderColor);
        Assert.AreEqual(XLColor.Automatic, border.TopBorderColor);
        Assert.AreEqual(XLColor.Automatic, border.BottomBorderColor);
        Assert.AreEqual(XLColor.Automatic, border.DiagonalBorderColor);
    }

    [Test]
    public void Border_lines_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Range("A1").AddConditionalFormat().Style;

        var returnedStyle = style.Border.SetLeftBorder(XLBorderStyleValues.Thin)
            .Border.SetLeftBorderColor(XLColor.Red)
            .Border.SetRightBorder(XLBorderStyleValues.Medium)
            .Border.SetRightBorderColor(XLColor.Green)
            .Border.SetTopBorder(XLBorderStyleValues.Thick)
            .Border.SetTopBorderColor(XLColor.Blue)
            .Border.SetBottomBorder(XLBorderStyleValues.Dashed)
            .Border.SetBottomBorderColor(XLColor.Yellow)
            .Border.SetDiagonalBorder(XLBorderStyleValues.Dotted)
            .Border.SetDiagonalBorderColor(XLColor.Black);

        Assert.AreSame(style, returnedStyle);
        var border = style.Border;
        Assert.AreEqual(XLBorderStyleValues.Thin, border.LeftBorder);
        Assert.AreEqual(XLColor.Red, border.LeftBorderColor);
        Assert.AreEqual(XLBorderStyleValues.Medium, border.RightBorder);
        Assert.AreEqual(XLColor.Green, border.RightBorderColor);
        Assert.AreEqual(XLBorderStyleValues.Thick, border.TopBorder);
        Assert.AreEqual(XLColor.Blue, border.TopBorderColor);
        Assert.AreEqual(XLBorderStyleValues.Dashed, border.BottomBorder);
        Assert.AreEqual(XLColor.Yellow, border.BottomBorderColor);
        Assert.AreEqual(XLBorderStyleValues.Dotted, border.DiagonalBorder);
        Assert.AreEqual(XLColor.Black, border.DiagonalBorderColor);
    }

    [Test]
    public void Border_outside_sets_all_four_lines_of_outline()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Border.SetOutsideBorder(XLBorderStyleValues.Double).Border.SetOutsideBorderColor(XLColor.Red);

        AssertFourLines(cf.Style.Border, XLBorderStyleValues.Double, XLColor.Red);
        Assert.True(((XLConditionalFormat)cf).FormatValue!.Border.Outline);
    }

    [Test]
    public void Border_inside_sets_all_four_lines_without_outline()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Border.SetInsideBorder(XLBorderStyleValues.Hair).Border.SetInsideBorderColor(XLColor.Blue);

        AssertFourLines(cf.Style.Border, XLBorderStyleValues.Hair, XLColor.Blue);
        Assert.False(((XLConditionalFormat)cf).FormatValue!.Border.Outline);
    }

    [Test]
    public void Border_can_be_set_by_assigning_cell_border()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Style.Border
            .SetLeftBorder(XLBorderStyleValues.Thin)
            .Border.SetTopBorder(XLBorderStyleValues.Thick)
            .Border.SetTopBorderColor(XLColor.Red)
            .Border.SetDiagonalBorder(XLBorderStyleValues.Dotted);
        var cf = ws.Range("A1").AddConditionalFormat();
        cf.Style.Border.SetInsideBorder(XLBorderStyleValues.Medium);

        cf.Style.Border = ws.Cell("B1").Style.Border;

        var border = cf.Style.Border;
        Assert.AreEqual(XLBorderStyleValues.Thin, border.LeftBorder);
        Assert.AreEqual(XLBorderStyleValues.None, border.RightBorder);
        Assert.AreEqual(XLBorderStyleValues.Thick, border.TopBorder);
        Assert.AreEqual(XLColor.Red, border.TopBorderColor);
        Assert.AreEqual(XLBorderStyleValues.None, border.BottomBorder);
        Assert.AreEqual(XLBorderStyleValues.Dotted, border.DiagonalBorder);
        Assert.True(((XLConditionalFormat)cf).FormatValue!.Border.Outline);
    }

    [Test]
    public void Border_can_be_set_by_assigning_border_of_other_conditional_format()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var source = ws.Range("A1").AddConditionalFormat();
        source.Style.Border
            .SetBottomBorder(XLBorderStyleValues.Thin)
            .Border.SetBottomBorderColor(XLColor.Green)
            .Border.SetDiagonalBorder(XLBorderStyleValues.Dashed)
            .Border.SetDiagonalUp();
        var target = ws.Range("B1").AddConditionalFormat();

        target.Style.Border = source.Style.Border;

        var border = target.Style.Border;
        Assert.AreEqual(XLBorderStyleValues.Thin, border.BottomBorder);
        Assert.AreEqual(XLColor.Green, border.BottomBorderColor);
        Assert.AreEqual(XLBorderStyleValues.Dashed, border.DiagonalBorder);
        Assert.True(border.DiagonalUp);
        Assert.False(border.DiagonalDown);
    }

    [Test]
    public void Border_diagonal_direction_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var style = ws.Range("A1").AddConditionalFormat().Style;
        Assert.False(style.Border.DiagonalUp);
        Assert.False(style.Border.DiagonalDown);

        var returnedStyle = style.Border.SetDiagonalUp().Border.SetDiagonalDown();

        Assert.AreSame(style, returnedStyle);
        Assert.True(style.Border.DiagonalUp);
        Assert.True(style.Border.DiagonalDown);

        style.Border.SetDiagonalUp(false).Border.SetDiagonalDown(false);

        Assert.False(style.Border.DiagonalUp);
        Assert.False(style.Border.DiagonalDown);
    }

    [Test]
    public void Border_equality_is_not_implemented()
    {
        // This test records a missing feature, not a requirement. If equality is added, change this test.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var border = ws.Range("A1").AddConditionalFormat().Style.Border;

        Assert.That(() => border.Equals(ws.Cell("A1").Style.Border), Throws.TypeOf<NotImplementedException>());
    }

    private static void AssertFourLines(IXLBorder border, XLBorderStyleValues expectedStyle, XLColor expectedColor)
    {
        Assert.AreEqual(expectedStyle, border.LeftBorder);
        Assert.AreEqual(expectedStyle, border.RightBorder);
        Assert.AreEqual(expectedStyle, border.TopBorder);
        Assert.AreEqual(expectedStyle, border.BottomBorder);
        Assert.AreEqual(expectedColor, border.LeftBorderColor);
        Assert.AreEqual(expectedColor, border.RightBorderColor);
        Assert.AreEqual(expectedColor, border.TopBorderColor);
        Assert.AreEqual(expectedColor, border.BottomBorderColor);
    }
}
