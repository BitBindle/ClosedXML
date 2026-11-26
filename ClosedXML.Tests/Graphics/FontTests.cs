using System;
using ClosedXML.Excel;
using ClosedXML.Graphics;
using NUnit.Framework;

namespace ClosedXML.Tests.Graphics
{
    [TestFixture]
    public class FontTests
    {
        private readonly IXLGraphicEngine _engine = DefaultGraphicEngine.Instance.Value;

        [TestCase]
        public void CalculatedTextWidth()
        {
            var textFont = new DummyFont("Calibri", 20);
            var textWidthPt = _engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont, 96);

            // The advances of the embedded font have a sum of 22920 font units. The font has 2048 units per em.
            Assert.That(textWidthPt, Is.EqualTo(22920d / 2048 * 20 / 72 * 96));
        }

        [Test]
        public void TextWidthIsSumOfCodePointWidths()
        {
            // The width has no rounding. Thus the sum of the widths of the code points is the width of the text.
            var textFont = new DummyFont("Calibri", 11);
            const string text = "Lorem ipsum, dolor (sit) amet 12.5 éß ΩЖ 日本";

            var textWidth = _engine.GetTextWidth(text, textFont, 96);
            var sumOfWidths = 0d;
            foreach (var codePoint in text.EnumerateRunes())
                sumOfWidths += _engine.GetTextWidth(codePoint.ToString(), textFont, 96);

            Assert.That(textWidth, Is.EqualTo(sumOfWidths).Within(1e-9));
        }

        [Test]
        public void TextWithLineBreakHasWidthOfLongestLine()
        {
            // A line break needs the text layout. The layout measures each line.
            var textFont = new DummyFont("Calibri", 11);

            var longLineFirst = _engine.GetTextWidth("Lorem ipsum dolor\nA", textFont, 96);
            var longLineLast = _engine.GetTextWidth("AB\nLorem ipsum dolor", textFont, 96);
            var longLine = _engine.GetTextWidth("Lorem ipsum dolor", textFont, 96);

            Assert.That(longLineFirst, Is.EqualTo(longLine).Within(1e-9));
            Assert.That(longLineLast, Is.EqualTo(longLine).Within(1e-9));
        }

        [Test]
        public void TextLayoutWidthHasNoRounding()
        {
            // A combining mark needs the text layout. The embedded font has no glyph positioning, so the width is the
            // sum of the advances: 13352 font units. The font has 2048 units per em.
            // An installed Calibri can position the mark, so the test uses only the embedded font.
            var engine = new DefaultGraphicEngine("NonExistentFallbackFont");
            var textFont = new DummyFont("SomeNonExistentFont", 12);

            var width = engine.GetTextWidth("Lorem ipsum e\u0301", textFont, 96);

            Assert.That(width, Is.EqualTo(13352d / 2048 * 12 / 72 * 96).Within(1e-9));
        }

        [TestCase]
        public void CalculatedTextHeight()
        {
            var textFont = new DummyFont("Calibri", 300);
            var textHeightPx = _engine.GetTextHeight(textFont, 96);
            Assert.That(textHeightPx, Is.EqualTo(500));
        }

        [TestCase]
        public void GetMaxDigitWidth()
        {
            var textFont = new DummyFont("Calibri", 11);
            var textWidthPx = _engine.GetMaxDigitWidth(textFont, 96);
            Assert.That(textWidthPx, Is.EqualTo(7.43359375d)); // Calibri,11 has a max digit width of 7 per spec 18.3.1.13
        }

        [TestCase]
        public void DescentIsPositive()
        {
            var textFont = new DummyFont("Calibri", 11);
            var textWidthPt = _engine.GetDescent(textFont, 96);
            Assert.That(textWidthPt, Is.EqualTo(3.666666666666667d));
        }

        [TestCase]
        public void NonExistentFontUsesFallback()
        {
            var nonExistentFont = new DummyFont("NonExistentFont", 100);
            var fallbackFont = new DummyFont("Microsoft Sans Serif", 100);

            var nonExistentFontWidth = _engine.GetTextWidth("ABCDEF text", nonExistentFont, 96);
            var fallbackFontWidth = _engine.GetTextWidth("ABCDEF text", fallbackFont, 96);
            Assert.That(nonExistentFontWidth, Is.EqualTo(fallbackFontWidth));

            var nonExistentFontHeight = _engine.GetTextHeight(nonExistentFont, 96);
            var fallbackFontHeight = _engine.GetTextHeight(fallbackFont, 96);
            Assert.That(nonExistentFontHeight, Is.EqualTo(fallbackFontHeight));
        }

        [Test]
        public void UseEmbeddedFontWhenFallbackFontIsNotPresent()
        {
            var nonExistentFont = new DummyFont("SomeNonExistentFont", 11);
            var engine = new DefaultGraphicEngine("NonExistentFallbackFont");
            Span<int> text = stackalloc int[1] { '8' };

            var box = engine.GetGlyphBox(text, nonExistentFont, new Dpi(96, 96));

            // Max digit width of CarlitoBare is 7, unlike MS Sans Serif which is 8
            Assert.AreEqual(7, box.AdvanceWidth);
        }

        [TestCase(true, false)]
        [TestCase(false, true)]
        [TestCase(true, true)]
        [Issue("137")]
        public void TextWidthUsesMetricsOfFontStyle(bool bold, bool italic)
        {
            // The embedded font has a file for each style, and the bold and italic glyphs have different widths.
            var engine = new DefaultGraphicEngine("NonExistentFallbackFont");
            var regularFont = new DummyFont("SomeNonExistentFont", 20);
            var styledFont = new DummyFont("SomeNonExistentFont", 20) { Bold = bold, Italic = italic };

            var regularWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", regularFont, 96);
            var styledWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", styledFont, 96);

            Assert.That(styledWidth, Is.Not.EqualTo(regularWidth));
        }

        [Test]
        [Issue("137")]
        public void BoldTextIsWiderThanRegularText()
        {
            var engine = new DefaultGraphicEngine("NonExistentFallbackFont");
            var regularFont = new DummyFont("SomeNonExistentFont", 20);
            var boldFont = new DummyFont("SomeNonExistentFont", 20) { Bold = true };

            var regularWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", regularFont, 96);
            var boldWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", boldFont, 96);

            Assert.That(boldWidth, Is.GreaterThan(regularWidth));
        }

        [Test]
        [Issue("137")]
        public void FontWithoutRequestedStyleUsesRegularStyle()
        {
            // TestFontA has only a regular style. The engine must use it for bold italic text without an exception.
            using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
            var engine = DefaultGraphicEngine.CreateOnlyWithFonts(fallbackFontStream);
            var regularFont = new DummyFont("Nonexistent Font", 20);
            var boldItalicFont = new DummyFont("Nonexistent Font", 20) { Bold = true, Italic = true };

            var regularWidth = engine.GetTextWidth("A", regularFont, 120);
            var boldItalicWidth = engine.GetTextWidth("A", boldItalicFont, 120);

            Assert.That(boldItalicWidth, Is.EqualTo(regularWidth));
        }

        [Test]
        [Issue("137")]
        public void FontWithoutBoldItalicStyleUsesBoldStyle()
        {
            // TestFontStyles has only a regular and a bold style.
            using var regularStream = TestHelper.GetStreamFromResource("Fonts.TestFontStyles-Regular.ttf");
            using var boldStream = TestHelper.GetStreamFromResource("Fonts.TestFontStyles-Bold.ttf");
            var engine = DefaultGraphicEngine.CreateOnlyWithFonts(regularStream, boldStream);
            var regularFont = new DummyFont("Nonexistent Font", 20);
            var boldFont = new DummyFont("Nonexistent Font", 20) { Bold = true };
            var boldItalicFont = new DummyFont("Nonexistent Font", 20) { Bold = true, Italic = true };

            var regularWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", regularFont, 96);
            var boldWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", boldFont, 96);
            var boldItalicWidth = engine.GetTextWidth("Lorem ipsum dolor sit amet", boldItalicFont, 96);

            Assert.That(boldItalicWidth, Is.EqualTo(boldWidth));
            Assert.That(boldItalicWidth, Is.Not.EqualTo(regularWidth));
        }

        [TestCase]
        public void CanSpecifyFallbackFontWithoutFileSystem()
        {
            using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
            var engine = DefaultGraphicEngine.CreateOnlyWithFonts(fallbackFontStream);

            var nonExistentFont = new DummyFont("Nonexistent Font", 20);
            var widthOfLetterA = engine.GetTextWidth("A", nonExistentFont, 120);

            // The advance of the letter A in TestFontA is 1886 font units. The font has 2048 units per em.
            const double expectedWidthOfLetterA = 1886d / 2048 * 20 / 72 * 120;
            Assert.AreEqual(expectedWidthOfLetterA, widthOfLetterA, 0.0001);
        }

        [TestCase]
        public void CanSpecifyExtraFontsAsStreamsWithoutFileSystem()
        {
            using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
            var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
            var engine = DefaultGraphicEngine.CreateOnlyWithFonts(fallbackFontStream, fontBStream);

            var widthOfLetterB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

            // The advance of the letter B in TestFontB is 602 font units. The font has 1000 units per em.
            const double expectedWidthOfLetterB = 602d / 1000 * 30 / 72 * 96;
            Assert.AreEqual(expectedWidthOfLetterB, widthOfLetterB, 0.0001);
        }

        [TestCase]
        public void Issue_1916_CanMeasureSpecificArabicText()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell(1, 1).Value = @"اصين";
            ws.Column(1).AdjustToContents();
        }

        private class DummyFont : IXLFontBase
        {
            public DummyFont(string name, double size)
            {
                FontName = name;
                FontSize = size;
            }

            public string FontName { get; set; }

            public double FontSize { get; set; }

            public bool Bold { get; set; }

            public bool Italic { get; set; }

            public bool Strikethrough { get; set; }

            public XLFontUnderlineValues Underline { get; set; } = XLFontUnderlineValues.None;

            public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

            public bool Shadow { get; set; }

            public XLColor FontColor { get; set; } = XLColor.Black;

            public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; } = XLFontFamilyNumberingValues.NotApplicable;

            public XLFontCharSet FontCharSet { get; set; } = XLFontCharSet.Default;

            public XLFontScheme FontScheme { get; set; }
        }
    }
}
