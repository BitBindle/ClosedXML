using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;

namespace ClosedXML.Tests.Utils;

/// <summary>
/// A graphic engine that returns known max digit widths of fonts (as measured by Excel on Windows at 96 DPI) and
/// delegates everything else to <see cref="DefaultGraphicEngine"/>. Column width is derived from max digit width
/// of the default font. The tests that load files with fonts that exist only on Windows (Arial, Yu Gothic) would
/// otherwise depend on the fonts installed on the machine.
/// </summary>
internal class FixedDigitWidthGraphicEngine : IXLGraphicEngine
{
    private readonly IXLGraphicEngine _engine = DefaultGraphicEngine.Instance.Value;
    private readonly Dictionary<(string Name, double Size), double> _maxDigitWidthsPx;

    /// <param name="maxDigitWidthsPx">Max digit width in pixels at 96 DPI for a font name and size.</param>
    public FixedDigitWidthGraphicEngine(Dictionary<(string Name, double Size), double> maxDigitWidthsPx)
    {
        _maxDigitWidthsPx = maxDigitWidthsPx;
    }

    public double GetMaxDigitWidth(IXLFontBase font, double dpiX)
    {
        if (_maxDigitWidthsPx.TryGetValue((font.FontName, font.FontSize), out var widthPx))
            return widthPx * dpiX / 96d;

        return _engine.GetMaxDigitWidth(font, dpiX);
    }

    public XLPictureInfo GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat)
        => _engine.GetPictureInfo(imageStream, expectedFormat);

    public double GetTextHeight(IXLFontBase font, double dpiY) => _engine.GetTextHeight(font, dpiY);

    public double GetTextWidth(string text, IXLFontBase font, double dpiX) => _engine.GetTextWidth(text, font, dpiX);

    public double GetDescent(IXLFontBase font, double dpiY) => _engine.GetDescent(font, dpiY);

    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        => _engine.GetGlyphBox(graphemeCluster, font, dpi);
}
