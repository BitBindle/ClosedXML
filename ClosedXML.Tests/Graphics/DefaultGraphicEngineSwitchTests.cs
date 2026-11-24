using System;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ClosedXML.Graphics;
using NUnit.Framework;

namespace ClosedXML.Tests.Graphics
{
    /// <summary>
    /// An application that never measures text can turn off the default graphic engine with a feature switch. Then
    /// the trimmer removes the engine, SixLabors.Fonts and the embedded fonts from a trimmed application.
    /// </summary>
    [TestFixture]
    [NonParallelizable]
    public class DefaultGraphicEngineSwitchTests
    {
        private const string SwitchName = "ClosedXML.Graphics.DefaultGraphicEngine.IsSupported";

        [TearDown]
        public void RestoreSwitch()
        {
            // A switch that is not set and a switch that is true give the same behavior.
            AppContext.SetSwitch(SwitchName, true);
        }

        [Test]
        public void Without_switch_workbook_measures_text_with_default_engine()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "A long text in a cell";

            ws.Column(1).AdjustToContents();

            Assert.That(ws.Column(1).Width, Is.GreaterThan(ws.ColumnWidth));
        }

        [Test]
        public void Switch_off_workbook_without_engine_throws_on_text_measure()
        {
            AppContext.SetSwitch(SwitchName, false);
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "A long text in a cell";

            Assert.Throws<NotSupportedException>(() => ws.Column(1).AdjustToContents());
        }

        [Test]
        public void Switch_off_workbook_without_engine_calculates_and_saves()
        {
            AppContext.SetSwitch(SwitchName, false);
            using var stream = new MemoryStream();
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = 5;
            ws.Cell("A2").FormulaA1 = "A1*2";

            Assert.AreEqual(XLCellValue.FromObject(10), ws.Cell("A2").Value);
            Assert.DoesNotThrow(() => wb.SaveAs(stream));
        }

        [Test]
        public void Switch_off_load_needs_engine_for_default_column_width()
        {
            // The load calculates the default column width from the max digit width of the normal font.
            using var stream = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                wb.AddWorksheet().Cell("A1").Value = 5;
                wb.SaveAs(stream);
            }

            AppContext.SetSwitch(SwitchName, false);

            Assert.Throws<NotSupportedException>(() => _ = new XLWorkbook(stream));
            using var loaded = new XLWorkbook(stream, new LoadOptions { GraphicEngine = new FixedGraphicEngine() });
            Assert.AreEqual(XLCellValue.FromObject(5), loaded.Worksheet(1).Cell("A1").Value);
        }

        [Test]
        public void Switch_off_default_engine_cannot_be_created()
        {
            AppContext.SetSwitch(SwitchName, false);

            Assert.Throws<NotSupportedException>(() => _ = new DefaultGraphicEngine("Arial"));
            Assert.Throws<NotSupportedException>(() => DefaultGraphicEngine.CreateOnlyWithFonts(new MemoryStream()));
        }

        [Test]
        public void Switch_off_workbook_measures_text_with_engine_of_load_options()
        {
            AppContext.SetSwitch(SwitchName, false);
            using var wb = new XLWorkbook(new LoadOptions { GraphicEngine = new FixedGraphicEngine() });
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "A long text in a cell";

            ws.Column(1).AdjustToContents();

            Assert.That(ws.Column(1).Width, Is.GreaterThan(ws.ColumnWidth));
        }

        [Test]
        public void Switch_off_workbook_measures_text_with_default_engine_of_load_options()
        {
            AppContext.SetSwitch(SwitchName, false);
            LoadOptions.DefaultGraphicEngine = new FixedGraphicEngine();
            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.AddWorksheet();
                ws.Cell("A1").Value = "A long text in a cell";

                ws.Column(1).AdjustToContents();

                Assert.That(ws.Column(1).Width, Is.GreaterThan(ws.ColumnWidth));
            }
            finally
            {
                LoadOptions.DefaultGraphicEngine = null;
            }
        }

        /// <summary>
        /// An engine without fonts: each glyph and each digit is 7 pixels wide at 96 DPI.
        /// </summary>
        private class FixedGraphicEngine : IXLGraphicEngine
        {
            public XLPictureInfo GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat) => throw new NotImplementedException();

            public double GetTextHeight(IXLFontBase font, double dpiY) => 20 * dpiY / 96;

            public double GetTextWidth(string text, IXLFontBase font, double dpiX) => text.Length * 7 * dpiX / 96;

            public double GetMaxDigitWidth(IXLFontBase font, double dpiX) => 7 * dpiX / 96;

            public double GetDescent(IXLFontBase font, double dpiY) => 4 * dpiY / 96;

            public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi) => new(7, 20, 4);
        }
    }
}
