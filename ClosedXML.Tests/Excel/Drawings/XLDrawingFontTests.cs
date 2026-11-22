using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Drawings;

/// <summary>
/// Test of <see cref="XLDrawingFont"/>.
/// </summary>
[TestFixture]
internal class XLDrawingFontTests
{
    [Test]
    public void New_font_has_default_values()
    {
        var font = new XLDrawingFont(new XLDrawingStyle());

        Assert.That(font.FontName, Is.EqualTo("Tahoma"));
        Assert.That(font.FontSize, Is.EqualTo(9));
        Assert.That(font.Underline, Is.EqualTo(XLFontUnderlineValues.None));
        Assert.That(font.FontColor, Is.EqualTo(XLColor.FromIndex(64)));
        Assert.That(font.Bold, Is.False);
        Assert.That(font.Italic, Is.False);
    }

    [Test]
    public void Font_properties_can_be_set()
    {
        var style = new XLDrawingStyle();
        var font = new XLDrawingFont(style);

        IXLDrawingStyle[] results =
        [
            font.SetBold(true),
            font.SetItalic(true),
            font.SetUnderline(XLFontUnderlineValues.Double),
            font.SetStrikethrough(true),
            font.SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Subscript),
            font.SetShadow(true),
            font.SetFontSize(12),
            font.SetFontColor(XLColor.Red),
            font.SetFontName("Arial"),
            font.SetFontFamilyNumbering(XLFontFamilyNumberingValues.Swiss),
            font.SetFontCharSet(XLFontCharSet.Greek),
            font.SetFontScheme(XLFontScheme.Minor),
        ];

        Assert.That(results, Is.All.SameAs(style));
        Assert.That(font.Bold, Is.True);
        Assert.That(font.Italic, Is.True);
        Assert.That(font.Underline, Is.EqualTo(XLFontUnderlineValues.Double));
        Assert.That(font.Strikethrough, Is.True);
        Assert.That(font.VerticalAlignment, Is.EqualTo(XLFontVerticalTextAlignmentValues.Subscript));
        Assert.That(font.Shadow, Is.True);
        Assert.That(font.FontSize, Is.EqualTo(12));
        Assert.That(font.FontColor, Is.EqualTo(XLColor.Red));
        Assert.That(font.FontName, Is.EqualTo("Arial"));
        Assert.That(font.FontFamilyNumbering, Is.EqualTo(XLFontFamilyNumberingValues.Swiss));
        Assert.That(font.FontCharSet, Is.EqualTo(XLFontCharSet.Greek));
        Assert.That(font.FontScheme, Is.EqualTo(XLFontScheme.Minor));
    }

    [Test]
    public void Font_flags_can_be_set_without_value()
    {
        var font = new XLDrawingFont(new XLDrawingStyle());

        font.SetBold();
        font.SetItalic();
        font.SetUnderline();
        font.SetStrikethrough();
        font.SetShadow();

        Assert.That(font.Bold, Is.True);
        Assert.That(font.Italic, Is.True);
        Assert.That(font.Underline, Is.EqualTo(XLFontUnderlineValues.Single));
        Assert.That(font.Strikethrough, Is.True);
        Assert.That(font.Shadow, Is.True);
    }
}
