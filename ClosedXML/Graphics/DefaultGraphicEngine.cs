using System;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SixLabors.Fonts;
using SixLabors.Fonts.Unicode;

namespace ClosedXML.Graphics
{
    public class DefaultGraphicEngine : IXLGraphicEngine
    {
        /// <summary>
        /// Carlito is a Calibri metric compatible font. This is a version stripped of everything but metric information
        /// to keep the embedded file small. It is reasonably accurate for many alphabets (contains 2531 glyphs). It has
        /// no glyph outlines, no TTF instructions, no substitutions, glyph positioning ect. It is created from Carlito
        /// font through strip-fonts.sh script.
        /// </summary>
        private const string EmbeddedFontName = "CarlitoBare";

        /// <summary>
        /// The name of the feature switch that turns off the engine. An application sets it to <c>false</c> with the
        /// MSBuild property <c>ClosedXMLDefaultGraphicEngineSupport</c>.
        /// </summary>
        internal const string IsSupportedSwitchName = "ClosedXML.Graphics.DefaultGraphicEngine.IsSupported";

        /// <summary>
        /// Vertical metrics (in 2048 units per em) of Calibri. Carlito has glyphs and advances of Calibri, but its
        /// vertical metrics are different (ascender 1950, descender -550), so the row heights would be different from
        /// Excel. When the embedded font is used as a fallback, use these metrics to get the heights of Calibri.
        /// </summary>
        private const int CalibriAscender = 1536;
        private const int CalibriDescender = -512;
        private const int CalibriUnitsPerEm = 2048;
        private const float FontMetricSize = 16f;
        private readonly ImageInfoReader[] _imageReaders =
        {
            new PngInfoReader(),
            new JpegInfoReader(),
            new GifInfoReader(),
            new TiffInfoReader(),
            new BmpInfoReader(),
            new EmfInfoReader(),
            new WmfInfoReader(),
            new WebpInfoReader(),
            new PcxInfoReader() // Due to poor magic detection, keep last
        };

        private readonly Lazy<IReadOnlyFontCollection> _fontCollection;
        private readonly string _fallbackFont;

        /// <summary>
        /// A font loaded font in the size <see cref="FontMetricSize"/>. There is no benefit in having multiple allocated instances, everything is just scaled at the moment.
        /// </summary>
        private readonly ConcurrentDictionary<MetricId, Font> _fonts = new();
        private readonly Func<MetricId, Font> _loadFont;

        /// <summary>
        /// Max digit width as a fraction of Em square. Multiply by font size to get pt size.
        /// </summary>
        private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();
        private readonly Func<MetricId, double> _calculateMaxDigitWidth;

        /// <summary>
        /// Can the application use the engine? The value is <c>true</c>, unless the application sets the feature switch
        /// <c>ClosedXML.Graphics.DefaultGraphicEngine.IsSupported</c> to <c>false</c>. Then a trimmed application does
        /// not contain the engine, SixLabors.Fonts and the embedded fonts.
        /// </summary>
        [FeatureSwitchDefinition(IsSupportedSwitchName)]
        internal static bool IsSupported => !AppContext.TryGetSwitch(IsSupportedSwitchName, out var isSupported) || isSupported;

        /// <summary>
        /// Get a singleton instance of the engine that uses <c>Microsoft Sans Serif</c> as a fallback font.
        /// </summary>
        public static Lazy<DefaultGraphicEngine> Instance { get; } = new(() => new("Microsoft Sans Serif"));

        /// <summary>
        /// Initialize a new instance of the engine.
        /// </summary>
        /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
        /// <exception cref="NotSupportedException">The application turned off the engine.</exception>
        public DefaultGraphicEngine(string fallbackFont)
        {
            if (!IsSupported)
                throw CreateNotSupportedException();

            if (string.IsNullOrWhiteSpace(fallbackFont))
                throw new ArgumentException(nameof(fallbackFont));

            var fontCollection = new FontCollection();
            AddEmbeddedFont(fontCollection);

            _fontCollection = new Lazy<IReadOnlyFontCollection>(() => fontCollection.AddSystemFonts());
            _fallbackFont = fallbackFont;
            _loadFont = LoadFont;
            _calculateMaxDigitWidth = CalculateMaxDigitWidth;
        }

        /// <summary>
        /// Initialize a new instance of the engine. The engine will be able to use system fonts and fonts loaded from external sources.
        /// </summary>
        /// <remarks>Useful/necessary for environments without an access to filesystem.</remarks>
        /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
        /// <param name="useSystemFonts">Should engine try to use system fonts? If false, system fonts won't be loaded which can significantly speed up library startup.</param>
        /// <param name="fontStreams">Extra fonts that should be loaded to the engine.</param>
        private DefaultGraphicEngine(Stream fallbackFontStream, bool useSystemFonts, Stream[] fontStreams)
        {
            if (!IsSupported)
                throw CreateNotSupportedException();

            if (fallbackFontStream is null)
                throw new ArgumentNullException(nameof(fallbackFontStream));

            if (fontStreams is null)
                throw new ArgumentNullException(nameof(fontStreams));

            var fontCollection = new FontCollection();
            AddEmbeddedFont(fontCollection);
            var fallbackFamily = fontCollection.Add(fallbackFontStream);
            foreach (var fontStream in fontStreams)
                fontCollection.Add(fontStream);

            _fontCollection = useSystemFonts
                ? new Lazy<IReadOnlyFontCollection>(() => fontCollection.AddSystemFonts())
                : new Lazy<IReadOnlyFontCollection>(() => fontCollection);
            _fallbackFont = fallbackFamily.Name;
            _loadFont = LoadFont;
            _calculateMaxDigitWidth = CalculateMaxDigitWidth;
        }

        /// <summary>
        /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
        /// It ignores all system fonts and that can lead to decrease of initialization time.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Font is determined by a name and style in the worksheet, but the font name must be mapped to a font file/stream.
        /// System fonts on Windows contain hundreds of font files that have to be checked to find the correct font
        /// file for the font name and style. That means to read hundreds of files and parse data inside them.
        /// Even though SixLabors.Fonts does this only once (lazily too) and stores data in a static variable, it is
        /// an overhead that can be avoided.
        /// </para>
        /// <para>
        /// This factory method is useful in several scenarios:
        /// <list type="bullet">
        ///   <item>Client side Blazor doesn't have access to any system fonts.</item>
        ///   <item>Worksheet contains only limited number of fonts. It might be sufficient to just load few fonts we are</item>
        /// </list>
        /// </para>
        /// </remarks>
        /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
        /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
        public static IXLGraphicEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams)
        {
            return new DefaultGraphicEngine(fallbackFontStream, false, fontStreams);
        }

        /// <summary>
        /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
        /// It also uses system fonts.
        /// </summary>
        /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
        /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
        public static IXLGraphicEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream, params Stream[] fontStreams)
        {
            return new DefaultGraphicEngine(fallbackFontStream, true, fontStreams);
        }

        public XLPictureInfo GetPictureInfo(Stream stream, XLPictureFormat expectedFormat)
        {
            foreach (var imageReader in _imageReaders)
            {
                if (imageReader.TryGetInfo(stream, out var dimensions))
                    return dimensions;
            }

            throw new ArgumentException("Unable to determine the format of the image.");
        }

        public double GetDescent(IXLFontBase font, double dpiY)
        {
            var metrics = GetMetrics(font);
            return GetDescent(font, dpiY, metrics);
        }

        private double GetDescent(IXLFontBase font, double dpiY, FontMetrics metrics)
        {
            var (_, descender) = GetVerticalMetrics(metrics);
            return PointsToPixels(-descender * font.FontSize / metrics.UnitsPerEm, dpiY);
        }

        public double GetMaxDigitWidth(IXLFontBase fontBase, double dpiX)
        {
            var metricId = new MetricId(fontBase);
            var maxDigitWidth = _maxDigitWidths.GetOrAdd(metricId, _calculateMaxDigitWidth);
            return PointsToPixels(maxDigitWidth * fontBase.FontSize, dpiX);
        }

        public double GetTextHeight(IXLFontBase font, double dpiY)
        {
            var metrics = GetMetrics(font);
            var (ascender, descender) = GetVerticalMetrics(metrics);
            return PointsToPixels((ascender - 2 * descender) * font.FontSize / metrics.UnitsPerEm, dpiY);
        }

        public double GetTextWidth(string text, IXLFontBase fontBase, double dpiX)
        {
            var font = GetFont(fontBase);
            var metrics = font.FontMetrics;
            if (!TryGetAdvanceSum(text, metrics, out var advanceFu))
                return GetLayoutTextWidth(text, font, fontBase, dpiX);

            return PointsToPixels(advanceFu * fontBase.FontSize / metrics.UnitsPerEm, dpiX);
        }

        /// <summary>
        /// Get the width of a text from a full text layout. The text layout is slow. It uses the substitutions of the
        /// font, for example the joined forms of Arabic letters. It measures each line of a text that has line breaks.
        /// </summary>
        private static double GetLayoutTextWidth(string text, Font font, IXLFontBase fontBase, double dpiX)
        {
            // The layout gives a width in whole pixels. At one pixel for each font unit, the layout does not round.
            var unitsPerEm = font.FontMetrics.UnitsPerEm;
            var dimensionsPx = TextMeasurer.MeasureAdvance(text, new TextOptions(new Font(font, unitsPerEm))
            {
                Dpi = 72, // Normalize DPI, so 1px is 1pt
                KerningMode = KerningMode.None
            });
            return PointsToPixels(dimensionsPx.Width / unitsPerEm * fontBase.FontSize, dpiX);
        }

        /// <summary>
        /// Add the advances of the code points of <paramref name="text"/>.
        /// </summary>
        /// <param name="text">The text to measure.</param>
        /// <param name="metrics">The metrics of the font.</param>
        /// <param name="advanceFu">The sum of the advances in font units.</param>
        /// <returns><c>false</c>, if the text has a code point that needs the text layout.</returns>
        private static bool TryGetAdvanceSum(string text, FontMetrics metrics, out long advanceFu)
        {
            advanceFu = 0;
            foreach (var rune in text.EnumerateRunes())
            {
                if (!HasContextFreeAdvance(rune))
                    return false;

                advanceFu += GetAdvance(metrics, rune.Value);
            }

            return true;
        }

        /// <summary>
        /// Is the advance of the code point independent of the code points around it? The ranges contain scripts that do
        /// not join or reorder letters. Other scripts, marks, control and format characters need the text layout.
        /// </summary>
        private static bool HasContextFreeAdvance(Rune rune)
        {
            var category = Rune.GetUnicodeCategory(rune);
            if (category is UnicodeCategory.Control
                or UnicodeCategory.Format
                or UnicodeCategory.NonSpacingMark
                or UnicodeCategory.SpacingCombiningMark
                or UnicodeCategory.EnclosingMark
                or UnicodeCategory.LineSeparator
                or UnicodeCategory.ParagraphSeparator)
                return false;

            return rune.Value is
                (>= 0x0020 and <= 0x02FF) or // Basic Latin to Spacing Modifier Letters
                (>= 0x0370 and <= 0x052F) or // Greek and Coptic, Cyrillic, Cyrillic Supplement
                (>= 0x1E00 and <= 0x1FFF) or // Latin Extended Additional, Greek Extended
                (>= 0x2000 and <= 0x2BFF) or // Punctuation, symbols, arrows, mathematical operators, box drawing
                (>= 0x3000 and <= 0x30FF) or // CJK Symbols and Punctuation, Hiragana, Katakana
                (>= 0x4E00 and <= 0x9FFF) or // CJK Unified Ideographs
                (>= 0xFF00 and <= 0xFFEF); // Halfwidth and Fullwidth Forms
        }

        /// <inheritdoc />
        public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        {
            // SixLabors.Fonts don't have a way to get a glyph representation of a cluster
            // without a TextRenderer that has unacceptable performance.
            var metric = GetMetrics(font);
            var advanceFu = 0;
            for (var i = 0; i < graphemeCluster.Length; ++i)
                advanceFu += GetAdvance(metric, graphemeCluster[i]);

            var emInPx = font.FontSize / 72d * dpi.X;
            var advancePx = PointsToPixels(advanceFu * font.FontSize / metric.UnitsPerEm, dpi.X);
            var descentPx = GetDescent(font, dpi.Y, metric);
            return new GlyphBox(
                (float)Math.Round(advancePx, MidpointRounding.AwayFromZero),
                (float)Math.Round(emInPx, MidpointRounding.AwayFromZero),
                (float)Math.Round(descentPx, MidpointRounding.AwayFromZero));
        }

        private FontMetrics GetMetrics(IXLFontBase fontBase)
        {
            var font = GetFont(fontBase);
            return font.FontMetrics;
        }

        /// <summary>
        /// Get ascender and descender in font units of <paramref name="metrics"/>.
        /// </summary>
        private static (double Ascender, double Descender) GetVerticalMetrics(FontMetrics metrics)
        {
            if (metrics.Description.FontFamilyInvariantCulture == EmbeddedFontName)
            {
                var scale = (double)metrics.UnitsPerEm / CalibriUnitsPerEm;
                return (CalibriAscender * scale, CalibriDescender * scale);
            }

            return (metrics.VerticalMetrics.Ascender, metrics.VerticalMetrics.Descender);
        }

        private Font GetFont(IXLFontBase fontBase)
        {
            return GetFont(new MetricId(fontBase));
        }

        private Font GetFont(MetricId metricId)
        {
            return _fonts.GetOrAdd(metricId, _loadFont);
        }

        private Font LoadFont(MetricId metricId)
        {
            // First try the specified fallback font. On windows, unknown fonts should use MS Sans Serif
            if (!_fontCollection.Value.TryGet(metricId.Name, out var fontFamily) &&
                !_fontCollection.Value.TryGet(_fallbackFont, out fontFamily))
            {
                // If not present, e.g. it's unlikely to be present on Linux, use embedded font as an ultimate fallback.
                fontFamily = _fontCollection.Value.Get(EmbeddedFontName);
            }

            // Size is irrelevant for metric. If the family does not have the style, SixLabors.Fonts uses the
            // nearest style that the family has, for example bold for bold italic, or regular.
            return fontFamily.CreateFont(FontMetricSize, metricId.Style);
        }

        /// <summary>
        /// The exception for a use of the engine when the application turned it off.
        /// </summary>
        internal static NotSupportedException CreateNotSupportedException()
        {
            return new NotSupportedException(
                $"The default graphic engine is not available, because the feature switch '{IsSupportedSwitchName}' is false. " +
                $"Set {nameof(LoadOptions)}.{nameof(LoadOptions.GraphicEngine)} or {nameof(LoadOptions)}.{nameof(LoadOptions.DefaultGraphicEngine)} to measure text or to read a picture.");
        }

        private void AddEmbeddedFont(FontCollection fontCollection)
        {
            var assembly = Assembly.GetExecutingAssembly();
            const string resourcePath = "ClosedXML.Graphics.Fonts.CarlitoBare-{0}.ttf";

            using var regular = assembly.GetManifestResourceStream(string.Format(resourcePath, "Regular"))!;
            fontCollection.Add(regular);

            using var bold = assembly.GetManifestResourceStream(string.Format(resourcePath, "Bold"))!;
            fontCollection.Add(bold);

            using var italic = assembly.GetManifestResourceStream(string.Format(resourcePath, "Italic"))!;
            fontCollection.Add(italic);

            using var boldItalic = assembly.GetManifestResourceStream(string.Format(resourcePath, "BoldItalic"))!;
            fontCollection.Add(boldItalic);
        }

        private double CalculateMaxDigitWidth(MetricId metricId)
        {
            var font = GetFont(metricId);
            var metrics = font.FontMetrics;
            var maxWidth = 0;
            for (var c = '0'; c <= '9'; ++c)
                maxWidth = Math.Max(maxWidth, GetAdvance(metrics, c));

            return maxWidth / (double)metrics.UnitsPerEm;
        }

        /// <summary>
        /// Get the advance of a code point in font units, without the substitutions of the font.
        /// </summary>
        private static int GetAdvance(FontMetrics metrics, int codePoint)
        {
            var containsMetrics = metrics.TryGetGlyphMetrics(
                new CodePoint(codePoint),
                TextAttributes.None,
                TextDecorations.None,
                LayoutMode.HorizontalTopBottom,
                ColorFontSupport.None,
                out var glyphs);

            // In SixLabors.Fonts 1.0.1, TryGetGlyphMetrics does not fail. For a missing glyph, it returns the .notdef
            // glyph 0. A later version can fail, so the code does a check of the result.
            if (!containsMetrics || glyphs is null)
                return 0;

            // An indexed loop does not allocate an enumerator for the interface.
            var advance = 0;
            for (var i = 0; i < glyphs.Count; ++i)
                advance += glyphs[i].AdvanceWidth;

            return advance;
        }

        private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;

        private readonly struct MetricId : IEquatable<MetricId>
        {
            public MetricId(IXLFontBase fontBase)
            {
                Name = fontBase.FontName;
                Style = GetFontStyle(fontBase);
            }

            public string Name { get; }

            public FontStyle Style { get; }

            public bool Equals(MetricId other) => Name == other.Name && Style == other.Style;

            public override bool Equals(object? obj) => obj is MetricId other && Equals(other);

            public override int GetHashCode() => (Name.GetHashCode() * 397) ^ (int)Style;

            private static FontStyle GetFontStyle(IXLFontBase fontBase)
            {
                return fontBase switch
                {
                    { Bold: true, Italic: true } => FontStyle.BoldItalic,
                    { Bold: true } => FontStyle.Bold,
                    { Italic: true } => FontStyle.Italic,
                    _ => FontStyle.Regular
                };
            }
        }
    }
}
