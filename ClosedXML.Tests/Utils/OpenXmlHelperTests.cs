using System;
using ClosedXML.Excel;
using ClosedXML.Utils;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;

namespace ClosedXML.Tests.Utils;

/// <summary>
/// Test of <see cref="OpenXmlHelper"/>.
/// </summary>
[TestFixture]
internal class OpenXmlHelperTests
{
    [Test]
    public void Number_format_without_source_is_not_changed()
    {
        using var wb = new XLWorkbook();
        var numberFormat = wb.AddWorksheet().Cell("A1").Style.NumberFormat;

        OpenXmlHelper.LoadNumberFormat(null!, numberFormat);

        Assert.That(numberFormat.NumberFormatId, Is.EqualTo(0));
    }

    [Test]
    public void Number_format_with_built_in_id_uses_id()
    {
        using var wb = new XLWorkbook();
        var numberFormat = wb.AddWorksheet().Cell("A1").Style.NumberFormat;
        var source = new NumberingFormat { NumberFormatId = 14, FormatCode = "0.00" };

        OpenXmlHelper.LoadNumberFormat(source, numberFormat);

        Assert.That(numberFormat.NumberFormatId, Is.EqualTo(14));
    }

    [TestCase(200u)]
    [TestCase(null)]
    public void Number_format_without_built_in_id_uses_format_code(uint? numberFormatId)
    {
        using var wb = new XLWorkbook();
        var numberFormat = wb.AddWorksheet().Cell("A1").Style.NumberFormat;
        var source = new NumberingFormat { FormatCode = "0.000" };
        if (numberFormatId is not null)
            source.NumberFormatId = numberFormatId.Value;

        OpenXmlHelper.LoadNumberFormat(source, numberFormat);

        Assert.That(numberFormat.Format, Is.EqualTo("0.000"));
    }

    [Test]
    public void Border_without_source_is_not_changed()
    {
        using var wb = new XLWorkbook();
        var border = wb.AddWorksheet().Cell("A1").Style.Border;

        OpenXmlHelper.LoadBorder(null!, border);

        Assert.That(border.LeftBorder, Is.EqualTo(XLBorderStyleValues.None));
        Assert.That(border.DiagonalBorder, Is.EqualTo(XLBorderStyleValues.None));
    }

    [Test]
    public void Border_loads_style_and_color_of_each_line()
    {
        using var wb = new XLWorkbook();
        var border = wb.AddWorksheet().Range("A1").AddConditionalFormat().Style.Border;
        var source = new Border
        {
            DiagonalUp = true,
            DiagonalDown = true,
            LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin, Color = new Color { Rgb = "FFFF0000" } },
            RightBorder = new RightBorder { Style = BorderStyleValues.Thick },
            TopBorder = new TopBorder { Color = new Color { Indexed = 10 } },
            BottomBorder = new BottomBorder { Style = BorderStyleValues.Dashed, Color = new Color { Theme = 4 } },
            DiagonalBorder = new DiagonalBorder { Style = BorderStyleValues.Double, Color = new Color { Theme = 5, Tint = 0.5 } },
        };

        OpenXmlHelper.LoadBorder(source, border);

        Assert.That(border.DiagonalUp, Is.True);
        Assert.That(border.DiagonalDown, Is.True);
        Assert.That(border.LeftBorder, Is.EqualTo(XLBorderStyleValues.Thin));
        Assert.That(border.LeftBorderColor, Is.EqualTo(XLColor.FromArgb(255, 255, 0, 0)));
        Assert.That(border.RightBorder, Is.EqualTo(XLBorderStyleValues.Thick));
        Assert.That(border.TopBorderColor, Is.EqualTo(XLColor.FromIndex(10)));
        Assert.That(border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Dashed));
        Assert.That(border.BottomBorderColor, Is.EqualTo(XLColor.FromTheme(XLThemeColor.Accent1)));
        Assert.That(border.DiagonalBorder, Is.EqualTo(XLBorderStyleValues.Double));
        Assert.That(border.DiagonalBorderColor, Is.EqualTo(XLColor.FromTheme(XLThemeColor.Accent2, 0.5)));
    }

    [Test]
    public void Border_without_diagonal_flags_keeps_diagonal_flags()
    {
        using var wb = new XLWorkbook();
        var border = wb.AddWorksheet().Cell("A1").Style.Border;
        border.DiagonalUp = true;

        OpenXmlHelper.LoadBorder(new Border(), border);

        Assert.That(border.DiagonalUp, Is.True);
        Assert.That(border.DiagonalDown, Is.False);
    }

    [Test]
    public void Fill_without_pattern_fill_is_not_changed()
    {
        using var wb = new XLWorkbook();
        var fill = wb.AddWorksheet().Cell("A1").Style.Fill;

        OpenXmlHelper.LoadFill(null!, fill, false);
        OpenXmlHelper.LoadFill(new Fill(), fill, false);

        Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.None));
    }

    [Test]
    public void Fill_without_pattern_type_is_solid()
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill
        {
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        });

        Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.Solid));
        Assert.That(fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(255, 0, 255, 0)));
    }

    [Test]
    public void Fill_with_none_pattern_has_no_pattern()
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill
        {
            PatternType = PatternValues.None,
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        });

        Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.None));
    }

    [Test]
    public void Differential_solid_fill_takes_background_color_from_background_color()
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" },
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        });

        Assert.That(fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(255, 0, 255, 0)));
    }

    [Test]
    public void Non_differential_solid_fill_takes_background_color_from_foreground_color()
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" },
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        }, false);

        Assert.That(fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(255, 255, 0, 0)));
    }

    [TestCase(true)]
    [TestCase(false)]
    public void Solid_fill_without_color_has_system_background_color(bool differential)
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill { PatternType = PatternValues.Solid }, differential);

        Assert.That(fill.BackgroundColor, Is.EqualTo(XLColor.FromIndex(64)));
    }

    [Test]
    public void Pattern_fill_takes_pattern_color_from_foreground_color()
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill
        {
            PatternType = PatternValues.DarkGray,
            ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" },
            BackgroundColor = new BackgroundColor { Rgb = "FF00FF00" },
        });

        Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.DarkGray));
        Assert.That(fill.PatternColor, Is.EqualTo(XLColor.FromArgb(255, 255, 0, 0)));
        Assert.That(fill.BackgroundColor, Is.EqualTo(XLColor.FromArgb(255, 0, 255, 0)));
    }

    [Test]
    public void Pattern_fill_without_colors_has_system_background_color()
    {
        using var wb = new XLWorkbook();
        var fill = LoadDifferentialFill(wb, new PatternFill { PatternType = PatternValues.DarkGray });

        Assert.That(fill.PatternType, Is.EqualTo(XLFillPatternValues.DarkGray));
        Assert.That(fill.PatternColor, Is.EqualTo(XLColor.Automatic));
        Assert.That(fill.BackgroundColor, Is.EqualTo(XLColor.FromIndex(64)));
    }

    [Test]
    public void Font_elements_without_value_use_default_values()
    {
        using var wb = new XLWorkbook();
        var richText = wb.AddWorksheet().Cell("A1").CreateRichText().AddText("A");
        var source = new RunProperties(
            new Bold(),
            new Italic { Val = false },
            new Underline(),
            new VerticalTextAlignment(),
            new FontScheme());
        richText.SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Superscript);
        richText.SetFontScheme(XLFontScheme.Major);

        OpenXmlHelper.LoadFont(source, richText);

        Assert.That(richText.Bold, Is.True);
        Assert.That(richText.Italic, Is.False);
        Assert.That(richText.Shadow, Is.False);
        Assert.That(richText.Underline, Is.EqualTo(XLFontUnderlineValues.Single));
        Assert.That(richText.VerticalAlignment, Is.EqualTo(XLFontVerticalTextAlignmentValues.Baseline));
        Assert.That(richText.FontScheme, Is.EqualTo(XLFontScheme.None));
    }

    [TestCase(0u, 0)]
    [TestCase(90u, 90)]
    [TestCase(135u, -45)]
    [TestCase(180u, -90)]
    [TestCase(255u, 255)]
    public void Rotation_above_90_degrees_is_rotation_below_horizontal(uint textRotation, int expected)
    {
        Assert.That(OpenXmlHelper.NormalizeRotation(textRotation), Is.EqualTo(expected));
    }

    [TestCase(181u)]
    [TestCase(254u)]
    [TestCase(256u)]
    public void Rotation_must_be_angle_or_vertical_text(uint textRotation)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => OpenXmlHelper.NormalizeRotation(textRotation));
    }

    [Test]
    public void Conversion_to_OpenXml_color_needs_color()
    {
        Assert.Throws<ArgumentNullException>(() => new Color().FromClosedXMLColor<Color>(null!));
        Assert.Throws<ArgumentNullException>(() => new DocumentFormat.OpenXml.Office2010.Excel.Color().FromClosedXMLColor<DocumentFormat.OpenXml.Office2010.Excel.Color>(null!));
    }

    private static IXLFill LoadDifferentialFill(XLWorkbook wb, PatternFill patternFill, bool differential = true)
    {
        var fill = wb.AddWorksheet().Range("A1").AddConditionalFormat().Style.Fill;
        OpenXmlHelper.LoadFill(new Fill(patternFill), fill, differential);
        return fill;
    }
}
