// Keep this file CodeMaid organized and cleaned
using System;
using System.Drawing;
using System.Threading;
using ClosedXML.Graphics;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A class that defines various aspects of a newly created workbook.
    /// </summary>
    public class LoadOptions
    {
        private System.Drawing.Point _dpi = new(96, 96);

        /// <summary>
        /// A graphics engine that will be used for workbooks without explicitly set engine.
        /// </summary>
        public static IXLGraphicEngine? DefaultGraphicEngine { internal get; set; }

        /// <summary>
        /// Should all formulas in a workbook be recalculated during load? Default value is <c>false</c>.
        /// </summary>
        public Boolean RecalculateAllFormulas { get; set; } = false;

        /// <summary>
        /// Graphic engine used by the workbook.
        /// </summary>
        public IXLGraphicEngine? GraphicEngine { get; set; }

        /// <summary>
        /// DPI for the workbook. Default is 96.
        /// </summary>
        /// <remarks>Used in various places, e.g. determining a physical size of an image without a DPI or to determine a size of a text in a cell.</remarks>
        public System.Drawing.Point Dpi
        {
            get => _dpi;
            set => _dpi = value.X > 0 && value.Y > 0 ? value : throw new ArgumentException("DPI must be positive");
        }

        /// <summary>
        /// <para>
        /// The option determines how should the XML parser handle XML attribute whose values do
        /// not match the expected XML simple type (e.g., attribute value is <c>one</c>, but it
        /// should be a <c>xsd:integer</c> or attribute value should be an enum <c>ST_BorderStyle</c>,
        /// but contains invalid value  <c>triangle</c> instead). If the option is enabled (default),
        /// the XML parser will throw when in encounters an attribute with a value that violates its
        /// simple type definition. <em><b>Disabling the option will hide these errors</b></em>.
        /// </para>
        /// <para>
        /// <list type="bullet">
        ///   <item>
        ///     <c>true</c> - the XML parser will throw when it encounters attribute with value
        ///     that doesn't match its simple type.
        ///   </item>
        ///   <item>
        ///     <c>false</c> - the XML parser will interpret attribute with invalid value as if
        ///     the attribute wasn't present.
        ///   </item>
        /// </list>
        /// Disabling this option can help in following cases: attribute is optional (malformed
        /// value is interpreted as missing) or attribute is optional, but has a default value
        /// (malformed value is interpreted as missing and a default value is used instead).
        /// </para>
        /// </summary>
        /// <remarks>
        /// This option only affects <c>ClosedXML.IO</c> parser. OOXML parts that are parsed by
        /// OpenXML SDK are not affected. Also, fix an OOXML producer that does this stuff.
        /// </remarks>
        public bool StrictAttributeParsing { get; set; } = true;

        /// <summary>
        /// <para>
        /// A function that tells if formulas can use a defined name or a table. The arguments are
        /// the sheet name of a sheet-scoped name and the name. The sheet name is <c>null</c> for a
        /// workbook-scoped name, for a table and for a name that the workbook doesn't have. The
        /// function returns <c>false</c> to block the name. The default is <c>null</c>: formulas
        /// can use each name.
        /// </para>
        /// <para>
        /// The function gets the stored name of a defined name or a table. For a name that the
        /// workbook doesn't have, it gets the text of the formula, so compare names without case.
        /// A table and a workbook-scoped defined name with the same name get the same arguments,
        /// so the function blocks both or neither.
        /// </para>
        /// <para>
        /// The evaluation of a blocked name throws an <see cref="InvalidOperationException"/>,
        /// also through <c>INDIRECT</c> and in a structured reference. A blocked name that the
        /// workbook doesn't have also throws.
        /// <see cref="IXLWorkbook.RecalculateAllFormulas(System.Collections.Generic.ICollection{IXLCell})"/>
        /// adds the cell to the failed cells. The calculation doesn't read the formula of a
        /// blocked defined name. A sheet-scoped name hides the workbook-scoped name with the same
        /// name on its sheet, also when the function blocks the sheet-scoped name.
        /// </para>
        /// </summary>
        /// <remarks>
        /// The calculation calls the function on its own thread. If more than one workbook uses
        /// the same function at the same time, the function must be thread-safe. The function
        /// must not change the workbook. The dependency tree of the workbook keeps the result for
        /// a defined name, so the function must give the same result for the same arguments each
        /// time.
        /// </remarks>
        public Func<string?, string, bool>? CalculationNameFilter { get; set; }

        /// <summary>
        /// A cancellation token to cancel recalculation. Recalculation may be triggered during
        /// various calls, e.g. <c>IXLCell.Value</c> can start recalculation of a workbook, but
        /// doesn't have a way to pass token.
        /// </summary>
        public CancellationToken CancellationToken { get; set; }

        /// <summary>
        /// A limit of how many bytes can be at most used by the calculation engine at any time
        /// during one calculation. Calculation means one run of calculation engine, it can be
        /// one cell or whole workbook. This is not a precise memory limit. It is used to check
        /// during allocation of memory buffers and other similar objects. It doesn't ensure that
        /// GC will collect buffers that are no longer used before new ones are allocated.
        /// The default is 0, which means no limit.
        /// </summary>
        public long MaxCalculationMemoryBytes
        {
            get;
            set => field = value >= 0
                ? value
                : throw new ArgumentOutOfRangeException(nameof(value), "Calculation memory limit must be non-negative.");
        } = 0;
    }
}
