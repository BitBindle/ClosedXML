using System.Collections;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.RichText;

/// <summary>
/// Tests of <see cref="XLPhonetics"/>.
/// </summary>
[TestFixture]
internal class XLPhoneticsTests
{
    [Test]
    public void New_phonetics_have_no_text_and_default_properties()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        var phonetics = ws.Cell("A1").CreateRichText().Phonetics;

        Assert.AreEqual(0, phonetics.Count);
        Assert.AreEqual(XLPhoneticType.FullWidthKatakana, phonetics.Type);
        Assert.AreEqual(XLPhoneticAlignment.Left, phonetics.Alignment);
        Assert.AreEqual(wb.Style.Font.FontName, phonetics.FontName);
        Assert.AreEqual(wb.Style.Font.FontSize, phonetics.FontSize);
        Assert.False(phonetics.Bold);
    }

    [Test]
    public void Font_properties_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var phonetics = ws.Cell("A1").CreateRichText().Phonetics;

        var returned = phonetics
            .SetBold()
            .SetItalic()
            .SetUnderline()
            .SetStrikethrough()
            .SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Superscript)
            .SetShadow()
            .SetFontSize(14)
            .SetFontColor(XLColor.Red)
            .SetFontName("Arial")
            .SetFontFamilyNumbering(XLFontFamilyNumberingValues.Swiss)
            .SetFontCharSet(XLFontCharSet.ShiftJIS)
            .SetFontScheme(XLFontScheme.Minor);

        Assert.AreSame(phonetics, returned);
        Assert.True(phonetics.Bold);
        Assert.True(phonetics.Italic);
        Assert.AreEqual(XLFontUnderlineValues.Single, phonetics.Underline);
        Assert.True(phonetics.Strikethrough);
        Assert.AreEqual(XLFontVerticalTextAlignmentValues.Superscript, phonetics.VerticalAlignment);
        Assert.True(phonetics.Shadow);
        Assert.AreEqual(14, phonetics.FontSize);
        Assert.AreEqual(XLColor.Red, phonetics.FontColor);
        Assert.AreEqual("Arial", phonetics.FontName);
        Assert.AreEqual(XLFontFamilyNumberingValues.Swiss, phonetics.FontFamilyNumbering);
        Assert.AreEqual(XLFontCharSet.ShiftJIS, phonetics.FontCharSet);
        Assert.AreEqual(XLFontScheme.Minor, phonetics.FontScheme);
    }

    [Test]
    public void Font_properties_can_be_set_to_specified_value()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var phonetics = ws.Cell("A1").CreateRichText().Phonetics;
        phonetics.SetBold().SetItalic().SetStrikethrough().SetShadow();

        phonetics
            .SetBold(false)
            .SetItalic(false)
            .SetUnderline(XLFontUnderlineValues.Double)
            .SetStrikethrough(false)
            .SetShadow(false);

        Assert.False(phonetics.Bold);
        Assert.False(phonetics.Italic);
        Assert.AreEqual(XLFontUnderlineValues.Double, phonetics.Underline);
        Assert.False(phonetics.Strikethrough);
        Assert.False(phonetics.Shadow);
    }

    [Test]
    public void Clear_font_sets_default_font()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var phonetics = ws.Cell("A1").CreateRichText().Phonetics;
        phonetics.SetBold().SetFontName("Arial");

        var returned = phonetics.ClearFont();

        Assert.AreSame(phonetics, returned);
        Assert.False(phonetics.Bold);
        Assert.AreEqual(wb.Style.Font.FontName, phonetics.FontName);
    }

    [Test]
    public void Alignment_and_type_can_be_set()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var phonetics = ws.Cell("A1").CreateRichText().Phonetics;

        var returned = phonetics.SetAlignment(XLPhoneticAlignment.Distributed).SetType(XLPhoneticType.Hiragana);

        Assert.AreSame(phonetics, returned);
        Assert.AreEqual(XLPhoneticAlignment.Distributed, phonetics.Alignment);
        Assert.AreEqual(XLPhoneticType.Hiragana, phonetics.Type);
    }

    [Test]
    public void Change_of_phonetics_changes_cell_value()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var richText = ws.Cell("A1").CreateRichText();
        richText.AddText("東京");

        richText.Phonetics.SetType(XLPhoneticType.Hiragana).SetAlignment(XLPhoneticAlignment.Center).SetBold().Add("とうきょう", 0, 2);

        var phonetics = ws.Cell("A1").GetRichText().Phonetics;
        Assert.AreEqual(XLPhoneticType.Hiragana, phonetics.Type);
        Assert.AreEqual(XLPhoneticAlignment.Center, phonetics.Alignment);
        Assert.True(phonetics.Bold);
        Assert.AreEqual("とうきょう", phonetics.Single().Text);
    }

    [Test]
    public void Clear_text_removes_all_phonetic_runs()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var richText = ws.Cell("A1").CreateRichText();
        richText.AddText("東京都");
        var phonetics = richText.Phonetics.Add("とう", 0, 1).Add("きょう", 1, 2);
        Assert.AreEqual(2, phonetics.Count);

        var returned = phonetics.ClearText();

        Assert.AreSame(phonetics, returned);
        Assert.AreEqual(0, phonetics.Count);
        Assert.IsEmpty(phonetics.ToList());
    }

    [Test]
    public void Phonetic_runs_are_enumerated_in_order_of_addition()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var richText = ws.Cell("A1").CreateRichText();
        richText.AddText("東京都");
        var phonetics = richText.Phonetics.Add("とう", 0, 1).Add("きょう", 1, 2);

        var genericTexts = phonetics.Select(p => (p.Text, p.Start, p.End)).ToArray();
        var nonGenericTexts = new ArrayList();
        foreach (IXLPhonetic phonetic in (IEnumerable)phonetics)
            nonGenericTexts.Add((phonetic.Text, phonetic.Start, phonetic.End));

        var expected = new[] { ("とう", 0, 1), ("きょう", 1, 2) };
        Assert.AreEqual(expected, genericTexts);
        Assert.AreEqual(expected, nonGenericTexts.ToArray());
    }

    [Test]
    public void Phonetics_with_same_content_are_equal()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var first = CreatePhonetics(ws.Cell("A1"));
        var second = CreatePhonetics(ws.Cell("A2"));

        Assert.True(first.Equals(first));
        Assert.True(first.Equals(second));
        Assert.False(first.Equals(null));
    }

    [Test]
    public void Phonetics_with_different_content_are_not_equal()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var phonetics = CreatePhonetics(ws.Cell("A1"));

        Assert.False(phonetics.Equals(CreatePhonetics(ws.Cell("A2")).ClearText()));
        Assert.False(phonetics.Equals(CreatePhonetics(ws.Cell("A3")).SetItalic()));
        Assert.False(phonetics.Equals(CreatePhonetics(ws.Cell("A4")).SetType(XLPhoneticType.HalfWidthKatakana)));
        Assert.False(phonetics.Equals(CreatePhonetics(ws.Cell("A5")).SetAlignment(XLPhoneticAlignment.Center)));
    }

    private static IXLPhonetics CreatePhonetics(IXLCell cell)
    {
        var richText = cell.CreateRichText();
        richText.AddText("東京");
        return richText.Phonetics.SetBold().Add("とうきょう", 0, 2);
    }
}
