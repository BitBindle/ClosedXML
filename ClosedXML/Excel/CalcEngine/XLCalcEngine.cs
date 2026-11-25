using ClosedXML.Excel.CalcEngine.Functions;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel.CalcEngine.Exceptions;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// CalcEngine parses strings and returns Expression objects that can
    /// be evaluated.
    /// </summary>
    /// <remarks>
    /// <para>This class has three extensibility points:</para>
    /// <para>Use the <b>RegisterFunction</b> method to define custom functions.</para>
    /// </remarks>
    internal class XLCalcEngine : ISheetListener, IWorkbookListener
    {
        private readonly CultureInfo _culture;
        private readonly FormulaParser _parser;
        private readonly CalculationVisitor _visitor;
        private DependencyTree? _dependencyTree;
        private XLCalculationChain? _chain;
        private readonly HashSet<string> _evaluatedNames = new();

        /// <param name="culture">The culture of the conversions in formulas.</param>
        /// <param name="wrongArgumentCountAsErrorValue">See <see cref="LoadOptions.WrongArgumentCountAsErrorValue"/>.</param>
        public XLCalcEngine(CultureInfo culture, bool wrongArgumentCountAsErrorValue = false)
        {
            _culture = culture;
            var funcRegistry = GetFunctionTable();
            _parser = new FormulaParser(funcRegistry, wrongArgumentCountAsErrorValue);
            _visitor = new CalculationVisitor(funcRegistry);
            _dependencyTree = null;
            _chain = null;
        }

        /// <summary>
        /// Parses a string into an <see cref="Formula"/>.
        /// </summary>
        /// <param name="expression">String to parse.</param>
        /// <returns>An formula that can be evaluated.</returns>
        public Formula Parse(string expression)
        {
            return _parser.GetAst(expression, isA1: true);
        }

        /// <summary>
        /// Add an array formula to the calc engine to manage dirty tracking and evaluation.
        /// </summary>
        internal void AddArrayFormula(Area range, XLCellFormula arrayFormula, XLWorksheet sheet)
        {
            if (_chain is not null && _dependencyTree is not null)
            {
                var area = new SheetArea(sheet.Name, range);
                _dependencyTree.AddFormula(area, arrayFormula, sheet.Workbook);
                _chain.AppendArea(area);
            }
        }

        /// <summary>
        /// Add a formula to the calc engine to manage dirty tracking and evaluation.
        /// </summary>
        internal void AddNormalFormula(SheetPoint point, string sheetName, XLCellFormula formula, XLWorkbook workbook)
        {
            if (_chain is not null && _dependencyTree is not null)
            {
                var pointArea = new SheetArea(sheetName, new Area(point.Point, point.Point));
                _dependencyTree.AddFormula(pointArea, formula, workbook);
                _chain.AddLast(point);
            }
        }

        /// <summary>
        /// Remove formula from dependency tree (=precedents won't mark
        /// it as dirty) and remove <paramref name="point"/> from the chain.
        /// Note that even if formula is used by many cells (e.g. array formula),
        /// it is fully removed from dependency tree, but each cells referencing
        /// the formula must be removed individually from calc chain.
        /// </summary>
        internal void RemoveFormula(SheetPoint point, XLCellFormula formula)
        {
            if (_chain is not null && _dependencyTree is not null)
            {
                _dependencyTree.RemoveFormula(formula);
                _chain.Remove(point);
            }
        }

        internal void OnAddedSheet(XLWorksheet sheet)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        internal void OnDeletingSheet(XLWorksheet sheet)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        internal void OnMovedSheet(XLWorksheet sheet)
        {
            // 3D references depend on the order of sheets.
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnInsertAreaAndShiftDown(XLWorksheet sheet, Area area)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnInsertAreaAndShiftRight(XLWorksheet sheet, Area area)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        public void OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea)
        {
            Purge(sheet.Workbook.WorksheetsInternal);
        }

        private void Purge(XLWorksheets sheets)
        {
            _dependencyTree = null;
            _chain = null;

            // Mark everything as dirty, because there can be stale values
            foreach (var sheet in sheets)
            {
                sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(Area.Full);
            }
        }

        internal void MarkDirty(XLWorksheet sheet, Point point)
        {
            MarkDirty(sheet, new Area(point, point));
        }

        internal void MarkDirty(XLWorksheet sheet, Area area)
        {
            if (_dependencyTree is not null)
            {
                var bookArea = new SheetArea(sheet.Name, area);
                _dependencyTree.MarkDirty(bookArea);
            }
        }

        /// <summary>
        /// Mark the formulas that read the <paramref name="area"/> as dirty. Unlike <see cref="MarkDirty(XLWorksheet, Area)"/>,
        /// build the dependency tree first if it doesn't exist.
        /// </summary>
        internal void MarkDependentsDirty(XLWorkbook wb, XLWorksheet sheet, Area area)
        {
            // An unparsable formula fails or throws when it is calculated, so the tree can skip it.
            EnsureChainAndTree(wb, skipUnparsable: true);
            MarkDirty(sheet, area);
        }

        [MemberNotNull(nameof(_chain), nameof(_dependencyTree))]
        private void EnsureChainAndTree(XLWorkbook wb, bool skipUnparsable)
        {
            // Lazy, so initialize chain from wb, if it is empty
            if (_chain is null || _dependencyTree is null)
            {
                _chain = XLCalculationChain.CreateFrom(wb);
                _dependencyTree = DependencyTree.CreateFrom(wb, skipUnparsable);
            }
        }

        /// <summary>
        /// Recalculate a workbook or a sheet.
        /// </summary>
        /// <param name="wb">The workbook.</param>
        /// <param name="recalculateSheetName">If set, only the formulas of this sheet are recalculated.</param>
        /// <param name="failed">
        /// If <c>null</c>, a cell that fails stops the recalculation with an exception. Otherwise a cell that is
        /// part of a cycle or whose formula throws keeps its value and is added to this collection. A cancellation
        /// still stops the recalculation.
        /// </param>
        /// <param name="records">
        /// If not <c>null</c>, each formula cell that uses an operation of a <see cref="XLCalculationRecordKind"/>
        /// is added to this collection, one time for each kind.
        /// </param>
        internal void Recalculate(XLWorkbook wb, string? recalculateSheetName, ICollection<IXLCell>? failed = null, ICollection<XLCalculationRecord>? records = null)
        {
            EnsureChainAndTree(wb, skipUnparsable: failed is not null);

            var sheetMap = wb.WorksheetsInternal
                .ToDictionary<XLWorksheet, string, (XLWorksheet Sheet, ValueSlice ValueSlice, FormulaSlice FormulaSlice)>(
                    sheet => sheet.Name,
                    sheet => (sheet, sheet.Internals.CellsCollection.ValueSlice, sheet.Internals.CellsCollection.FormulaSlice),
                    XLHelper.SheetComparer);

            // A recalculation without records also replaces the recorder of the thread. Else a recalculation that
            // runs during another one (e.g. from code in the collection of records) reports into its current cell.
            var outerRecorder = CalculationRecorder.Current;
            var recorder = records is null ? null : new CalculationRecorder(records);
            CalculationRecorder.Current = recorder;
            try
            {
                // Each outer loop moves chain one cell ahead.
                while (_chain.MoveAhead())
                {
                    // Inner loop that pushes supporting formulas ahead of current.
                    // It ends when a cell has been calculated and thus chain can move ahead.
                    while (true)
                    {
                        var current = _chain.Current;
                        var sheetName = current.SheetName;

                        // Skip dirty cells from sheets that are not being recalculated
                        if (recalculateSheetName is not null && !XLHelper.SheetComparer.Equals(sheetName, recalculateSheetName))
                        {
                            // Even though cell is dirty, it's in the ignored sheet and
                            // thus chain can move ahead.
                            break;
                        }

                        if (!sheetMap.TryGetValue(sheetName, out var sheetInfo))
                        {
                            throw new InvalidOperationException($"Sheet of a calc chain cell {current} doesn't exist in the workbook.");
                        }

                        var cellFormula = sheetInfo.FormulaSlice.Get(current.Point);
                        if (cellFormula is null)
                        {
                            throw new InvalidOperationException($"Calculation chain contains a {current}, but the cell doesn't contain a formula.");
                        }

                        if (_chain.IsCurrentInCycle)
                        {
                            if (failed is null)
                                throw new InvalidOperationException($"Formula in a cell {current} is part of a cycle.");

                            AddFailedCell(failed, cellFormula, sheetInfo.Sheet, current.Point);
                            break;
                        }

                        if (!cellFormula.IsDirty)
                            break;

                        try
                        {
                            // The records of an array formula go to its first cell, the cell with the formula.
                            recorder?.BeginCell(sheetInfo.Sheet, cellFormula.Type == FormulaType.Array ? cellFormula.Range.FirstPoint : current.Point);
                            ApplyFormula(cellFormula, current.Point, sheetInfo.Sheet, sheetInfo.ValueSlice,
                                recalculateSheetName);
                            cellFormula.IsDirty = false;
                        }
                        catch (GettingDataException ex)
                        {
                            // The engine evaluates the cell again after the dirty cell.
                            recorder?.DiscardCell();
                            _chain.MoveToCurrent(ex.Point);
                            continue;
                        }
                        catch (Exception ex) when (failed is not null && ex is not (OperationCanceledException or OutOfMemoryException))
                        {
                            // The records go to the collection after the cell is no longer dirty. An exception
                            // of a collection is not a failure of the formula, so it stops the recalculation.
                            var failedKinds = recorder?.TakeCell() ?? default;
                            AddFailedCell(failed, cellFormula, sheetInfo.Sheet, current.Point);
                            recorder?.Write(failedKinds);

                            break;
                        }

                        if (recorder is not null)
                            recorder.Write(recorder.TakeCell());

                        // Break out of the inner loop, a dirty cell has been
                        // calculated and thus chain can move ahead.
                        break;
                    }
                }
            }
            finally
            {
                CalculationRecorder.Current = outerRecorder;

                // Super important to clean up the chain for next recalculation.
                // Chain contains shared data and not cleaning it would cause hard
                // to diagnose issues.
                _chain.Reset();
            }
        }

        private static void AddFailedCell(ICollection<IXLCell> failed, XLCellFormula formula, XLWorksheet sheet, Point point)
        {
            // The cell keeps its value. Its dependents use that value, because the cell is no longer dirty.
            formula.IsDirty = false;
            failed.Add(sheet.Cell(point.Row, point.Column));
        }

        private void ApplyFormula(XLCellFormula formula, Point appliedPoint, XLWorksheet sheet, ValueSlice valueSlice, string? recalculateSheetName)
        {
            var formulaText = formula.A1;
            if (formula.Type == FormulaType.Normal)
            {
                var single = EvaluateFormula(
                    formulaText,
                    sheet.Workbook,
                    sheet,
                    new XLAddress(sheet, appliedPoint.Row, appliedPoint.Column, true, true),
                    recalculateSheetName: recalculateSheetName);
                valueSlice.SetCellValue(appliedPoint, single.ToCellValue());
            }
            else if (formula.Type == FormulaType.Array)
            {
                // The point can be any point in an array, so we can't use it.
                var range = formula.Range;
                var leftTopCorner = range.FirstPoint;
                var masterCell = sheet.Cell(leftTopCorner.Row, leftTopCorner.Column);
                var array = EvaluateArrayFormula(formulaText, masterCell, recalculateSheetName);

                // The array from formula can be smaller or larger than the
                // range of cells it should fit into. Broadcast it to the size.
                var result = array.Broadcast(range.Height, range.Width);

                // Copy value to the value slice
                for (var rowIdx = 0; rowIdx < result.Height; ++rowIdx)
                {
                    for (var colIdx = 0; colIdx < result.Width; ++colIdx)
                    {
                        var cellValue = result[rowIdx, colIdx];
                        var row = range.FirstPoint.Row + rowIdx;
                        var column = range.FirstPoint.Column + colIdx;
                        valueSlice.SetCellValue(new Point(row, column), cellValue.ToCellValue());
                    }
                }
            }
            else
            {
                throw new NotImplementedException($"Evaluation of formula type '{formula.Type}' is not supported.");
            }
        }

        /// <summary>
        /// Evaluates a normal formula.
        /// </summary>
        /// <param name="expression">Expression to evaluate.</param>
        /// <param name="wb">Workbook where is formula being evaluated.</param>
        /// <param name="ws">Worksheet where is formula being evaluated.</param>
        /// <param name="address">Address of formula.</param>
        /// <param name="recursive">Should the data necessary for this formula (not deeper ones)
        /// be calculated recursively? Used only for non-cell calculations.</param>
        /// <param name="recalculateSheetName">
        /// If set, calculation  will allow dirty reads from other sheets than the passed one.
        /// </param>
        /// <returns>The value of the expression.</returns>
        /// <remarks>
        /// If you are going to evaluate the same expression several times,
        /// it is more efficient to parse it only once using the <see cref="Parse"/>
        /// method and then using the Expression.Evaluate method to evaluate
        /// the parsed expression.
        /// </remarks>
        internal ScalarValue EvaluateFormula(string expression, XLWorkbook? wb = null, XLWorksheet? ws = null, IXLAddress? address = null, bool recursive = false, string? recalculateSheetName = null)
        {
            var ctx = new CalcContext(this, _culture, wb, ws, address, recursive)
            {
                RecalculateSheetName = recalculateSheetName
            };
            var result = EvaluateFormula(expression, ctx);
            if (ctx.UseImplicitIntersection)
            {
                result = result.Match(
                    () => AnyValue.Blank,
                    logical => logical,
                    number => number,
                    text => text,
                    error => error,
                    array => array[0, 0].ToAnyValue(),
                    reference => reference);
            }

            return ToCellContentValue(result, ctx);
        }

        /// <summary>
        /// Evaluates a formula without implicit intersection of the result. A reference stays a reference, and an
        /// array stays an array. As <see cref="EvaluateFormula(string, XLWorkbook?, XLWorksheet?, IXLAddress?, bool, string?)"/>
        /// with <c>recursive</c>, the engine calculates a dirty cell that the formula reads first.
        /// </summary>
        internal AnyValue EvaluateFormulaWithoutIntersection(string expression, XLWorksheet ws, IXLAddress? address)
        {
            var ctx = new CalcContext(this, _culture, ws.Workbook, ws, address, recursive: true);
            return EvaluateFormula(expression, ctx);
        }

        private Array EvaluateArrayFormula(string expression, XLCell masterCell, string? recalculateSheetName)
        {
            var ctx = new CalcContext(this, _culture, masterCell)
            {
                IsArrayCalculation = true,
                RecalculateSheetName = recalculateSheetName
            };
            var result = EvaluateFormula(expression, ctx);
            if (result.TryPickSingleOrMultiValue(out var single, out var multi, ctx))
                return new ScalarArray(single, 1, 1);

            return multi;
        }

        internal AnyValue EvaluateName(string nameFormula, XLWorksheet ws)
        {
            // A name can refer to itself, e.g. through INDIRECT. The evaluation of a name formula
            // always gives same result, so a repeated formula is a cycle that would overflow the stack.
            if (!_evaluatedNames.Add(nameFormula))
                throw new InvalidOperationException($"Name formula '{nameFormula}' is part of a cycle.");

            try
            {
                var ctx = new CalcContext(this, _culture, ws.Workbook, ws, null);
                return EvaluateFormula(nameFormula, ctx);
            }
            finally
            {
                _evaluatedNames.Remove(nameFormula);
            }
        }

        /// <summary>
        /// Evaluate a text of a reference, e.g. for <c>INDIRECT</c>. The text can be a reference
        /// in A1 or R1C1 style, a structured reference or a name. Anything that isn't
        /// a reference is <c>#REF!</c>.
        /// </summary>
        internal AnyValue EvaluateReferenceText(string text, bool isA1, CalcContext ctx)
        {
            ValueNode root;
            try
            {
                root = _parser.GetAst(text, isA1).AstRoot;
            }
            catch (ExpressionParseException)
            {
                return XLError.CellReference;
            }

            if (root is ReferenceNode referenceNode)
            {
                XLWorksheet? sheet = null;
                var prefix = referenceNode.Prefix;
                if (prefix is not null)
                {
                    if (prefix.File is not null || prefix.Sheet is null || !ctx.Workbook.TryGetWorksheet(prefix.Sheet, out sheet))
                        return XLError.CellReference;
                }

                var area = referenceNode.ReferenceArea;
                var isRelative = !isA1 && (
                    area.First.RowType == ReferenceAxisType.Relative || area.First.ColumnType == ReferenceAxisType.Relative ||
                    area.Second.RowType == ReferenceAxisType.Relative || area.Second.ColumnType == ReferenceAxisType.Relative);

                // A relative R1C1 reference needs a formula cell, but a name formula has none.
                if (isRelative && !ctx.HasFormulaAddress)
                    return XLError.CellReference;

                var anchor = isRelative ? ctx.FormulaPoint : default;
                return new Reference(new XLReference(area).ToRangeAddress(sheet, anchor));
            }

            if (root is not (NameNode or StructuredReferenceNode))
                return XLError.CellReference;

            var value = root.Accept(ctx, _visitor);
            return value.IsReference ? value : XLError.CellReference;
        }

        private AnyValue EvaluateFormula(string expression, CalcContext ctx)
        {
            var formula = Parse(expression);
            var result = formula.AstRoot.Accept(ctx, _visitor);
            return result;
        }

        // build/get static keyword table
        private FunctionRegistry GetFunctionTable()
        {
            var fr = new FunctionRegistry();

            // register built-in functions (and constants)
            Engineering.Register(fr);
            Information.Register(fr);
            Logical.Register(fr);
            Lookup.Register(fr);
            MathTrig.Register(fr);
            Text.Register(fr);
            Statistical.Register(fr);
            DateAndTime.Register(fr);
            Financial.Register(fr);

            return fr;
        }

        /// <summary>
        /// Convert any kind of formula value to value returned as a content of a cell.
        /// <list type="bullet">
        ///    <item><c>bool</c> - represents a logical value.</item>
        ///    <item><c>double</c> - represents a number and also date/time as serial date-time.</item>
        ///    <item><c>string</c> - represents a text value.</item>
        ///    <item><see cref="XLError" /> - represents a formula calculation error.</item>
        /// </list>
        /// </summary>
        private static ScalarValue ToCellContentValue(AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out var scalar, out var collection))
                return scalar;

            if (collection.TryPickT0(out var array, out var reference))
            {
                return array![0, 0];
            }

            if (reference!.TryGetSingleCellValue(out var cellValue, ctx))
                return cellValue;

            var intersected = reference.ImplicitIntersection(ctx.FormulaAddress);
            if (!intersected.TryPickT0(out var singleCellReference, out var error))
                return error;

            if (!singleCellReference!.TryGetSingleCellValue(out var singleCellValue, ctx))
                throw new InvalidOperationException("Got multi cell reference instead of single cell reference.");

            return singleCellValue;
        }

        void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
        {
            if (_dependencyTree is not null)
                _dependencyTree.RenameSheet(oldSheetName, newSheetName);
        }
    }

    internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue> arg);
}
