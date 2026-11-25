using System;
using System.Collections.Generic;
using ClosedXML.Extensions;
using ClosedXML.Parser;

namespace ClosedXML.Excel.CalcEngine
{
    internal class FormulaParser
    {
        private readonly AstFactory _nodeFactoryA1;
        private readonly AstFactory _nodeFactoryR1C1;
        private readonly AstFactory? _uncheckedFactoryA1;
        private readonly AstFactory? _uncheckedFactoryR1C1;

        /// <param name="functionRegistry">The functions that formulas can call.</param>
        /// <param name="wrongArgumentCountAsErrorValue">
        /// <c>true</c> - a formula with a wrong number of arguments for a function is <c>#VALUE!</c>.
        /// <c>false</c> - the parse of such a formula throws an <see cref="ExpressionParseException"/>.
        /// </param>
        public FormulaParser(FunctionRegistry functionRegistry, bool wrongArgumentCountAsErrorValue = false)
        {
            var argumentCount = wrongArgumentCountAsErrorValue ? ArgumentCount.Signal : ArgumentCount.Throw;
            _nodeFactoryA1 = new AstFactory(functionRegistry, true, argumentCount);
            _nodeFactoryR1C1 = new AstFactory(functionRegistry, false, argumentCount);
            if (wrongArgumentCountAsErrorValue)
            {
                _uncheckedFactoryA1 = new AstFactory(functionRegistry, true, ArgumentCount.Ignore);
                _uncheckedFactoryR1C1 = new AstFactory(functionRegistry, false, ArgumentCount.Ignore);
            }
        }

        /// <summary>
        /// Parse a formula into an abstract syntax tree.
        /// </summary>
        public Formula GetAst(string formula, bool isA1)
        {
            // Equality sign at the beginning of formula is only visualization in the GUI, real formulas don't have it.
            if (formula.Length > 0 && formula[0] == '=')
                formula = formula.Substring(1);

            try
            {
                return new Formula(formula, Parse(formula, isA1, isA1 ? _nodeFactoryA1 : _nodeFactoryR1C1));
            }
            catch (ParsingException ex)
            {
                throw new ExpressionParseException(ex.Message);
            }
            catch (WrongArgumentCountException)
            {
                // The factory stops at the first wrong call. Parse the rest of the formula too, so that
                // a syntax error after the call is still a parse error.
                try
                {
                    Parse(formula, isA1, isA1 ? _uncheckedFactoryA1! : _uncheckedFactoryR1C1!);
                }
                catch (ParsingException ex)
                {
                    throw new ExpressionParseException(ex.Message);
                }

                // The whole formula is the error, so IFERROR in the formula doesn't hide it.
                return new Formula(formula, new ScalarNode(XLError.IncompatibleValue));
            }
        }

        private static ValueNode Parse(string formula, bool isA1, AstFactory factory)
        {
            return isA1
                ? FormulaParser<ScalarValue, ValueNode, string>.CellFormulaA1(formula, formula, factory)
                : FormulaParser<ScalarValue, ValueNode, string>.CellFormulaR1C1(formula, formula, factory);
        }

        /// <summary>
        /// What the factory does with a call of a known function with a wrong number of arguments.
        /// </summary>
        private enum ArgumentCount
        {
            /// <summary>Throw an <see cref="ExpressionParseException"/>.</summary>
            Throw,

            /// <summary>Throw a <see cref="WrongArgumentCountException"/>.</summary>
            Signal,

            /// <summary>Create the function node.</summary>
            Ignore,
        }

        /// <summary>
        /// The factory throws this exception for a wrong number of arguments, when the formula is
        /// <c>#VALUE!</c> instead of a parse error.
        /// </summary>
        private sealed class WrongArgumentCountException : Exception
        {
        }

        /// <summary>
        /// Factory to create abstract syntax tree for a formula in A1 notation.
        /// </summary>
        private sealed class AstFactory : IAstFactory<ScalarValue, ValueNode, string>
        {
            /// <summary>
            /// A prefix for so-called future functions. Excel can add functions, but to avoid name collisions,
            /// it prefixes names of function with this prefix. The prefix is omitted from GUI.
            /// </summary>
            /// <example>
            /// If you write <c>CONCAT(A1,B1)</c> in Excel 2021 (not present in Excel 2013), it is saved to the
            /// worksheet file as <c>_xlfn.CONCAT(A1,B1)</c>, but the Excel GUI will show only <c>CONCAT(A1,B1)</c>,
            /// without the <c>_xlfn</c>.
            /// </example>
            private const string DefaultFunctionNameSpace = "_xlfn";

            private readonly FunctionRegistry _functionRegistry;
            private readonly bool _isA1;
            private readonly ArgumentCount _argumentCount;

            internal AstFactory(FunctionRegistry functionRegistry, bool isA1, ArgumentCount argumentCount)
            {
                _functionRegistry = functionRegistry;
                _isA1 = isA1;
                _argumentCount = argumentCount;
            }

            public ScalarValue LogicalValue(string context, SymbolRange range, bool logical) => logical;

            public ScalarValue NumberValue(string context, SymbolRange range, double number) => number;

            public ScalarValue TextValue(string context, SymbolRange range, ReadOnlySpan<char> text) => text.ToString();

            public ScalarValue ErrorValue(string context, SymbolRange range, ReadOnlySpan<char> errorText)
            {
                return GetErrorValue(errorText);
            }

            public ValueNode ArrayNode(string context, SymbolRange range, int rows, int columns,
                IReadOnlyList<ScalarValue> elements)
            {
                var array = new LiteralArray(rows, columns, elements);
                return new ArrayNode(array);
            }

            public ValueNode BlankNode(string context, SymbolRange range)
            {
                return new ScalarNode(ScalarValue.Blank);
            }

            public ValueNode LogicalNode(string context, SymbolRange range, bool logical)
            {
                return new ScalarNode(logical);
            }

            public ValueNode ErrorNode(string context, SymbolRange range, ReadOnlySpan<char> errorText)
            {
                var error = GetErrorValue(errorText);
                return new ScalarNode(error);
            }

            public ValueNode NumberNode(string context, SymbolRange range, double number)
            {
                return new ScalarNode(number);
            }

            public ValueNode TextNode(string context, SymbolRange range, ReadOnlySpan<char> text)
            {
                return new ScalarNode(text.ToString());
            }

            public ValueNode Reference(string context, SymbolRange range, ReferenceArea area)
            {
                return new ReferenceNode(null, area, _isA1);
            }

            public ValueNode SheetReference(string context, SymbolRange range, ReadOnlySpan<char> sheet, ReferenceArea area)
            {
                var prefixNode = new PrefixNode(null, sheet.ToString(), null, null);
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode BangReference(string context, SymbolRange range, ReferenceArea reference)
            {
                return new NotSupportedNode("Bang reference");
            }

            public ValueNode Reference3D(string context, SymbolRange range, ReadOnlySpan<char> firstSheet, ReadOnlySpan<char> lastSheet,
                ReferenceArea area)
            {
                var prefixNode = new PrefixNode(null, null, firstSheet.ToString(), lastSheet.ToString());
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode ExternalSheetReference(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> sheet,
                ReferenceArea area)
            {
                var fileNode = new FileNode(workbookIndex);
                var prefixNode = new PrefixNode(fileNode, sheet.ToString(), null, null);
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode ExternalReference3D(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> firstSheet,
                ReadOnlySpan<char> lastSheet, ReferenceArea area)
            {
                var fileNode = new FileNode(workbookIndex);
                var prefixNode = new PrefixNode(fileNode, null, firstSheet.ToString(), lastSheet.ToString());
                return new ReferenceNode(prefixNode, area, _isA1);
            }

            public ValueNode Function(string context, SymbolRange range, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var functionName = name.ToString();
                return GetFunctionNode(null, functionName, args);
            }

            public ValueNode Function(string context, SymbolRange range, ReadOnlySpan<char> sheetName, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(null, sheetName.ToString(), null, null);
                return GetFunctionNode(prefixNode, name.ToString(), args);
            }

            public ValueNode ExternalFunction(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> sheet,
                ReadOnlySpan<char> name, IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet.ToString(), null, null);
                return GetFunctionNode(prefixNode, name.ToString(), args);
            }

            public ValueNode ExternalFunction(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> name,
                IReadOnlyList<ValueNode> args)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
                return GetFunctionNode(prefixNode, name.ToString(), args);
            }

            public ValueNode CellFunction(string context, SymbolRange range, RowCol cell,
                IReadOnlyList<ValueNode> args)
            {
                // Grammar technically allows to evaluate a function from a different cell. The intended
                // usage is likely for lambda functions. Excel (as of 2022) doesn't do that, so use preference
                // as LOG10. Parser doesn't know about names of functions, so names such as LOG10 will always end up
                // here.
                var functionName = context.Substring(range.Start, context.IndexOf('(', range.Start) - range.Start);
                if (_functionRegistry.TryGetFunc(functionName, out _, out _))
                    return new FunctionNode(functionName, args);

                // Nonexistent function is evaluated to #NAME?, but cell function should be evaluated to #REF!
                return new ScalarNode(XLError.CellReference);
            }

            public ValueNode StructureReference(string context, SymbolRange range, StructuredReferenceArea area,
                ReadOnlySpan<char> firstColumn, ReadOnlySpan<char> lastColumn)
            {
                return new StructuredReferenceNode(null, null, area, ColumnOrNull(firstColumn), ColumnOrNull(lastColumn));
            }

            public ValueNode StructureReference(string context, SymbolRange range, ReadOnlySpan<char> table, StructuredReferenceArea area,
                ReadOnlySpan<char> firstColumn, ReadOnlySpan<char> lastColumn)
            {
                return new StructuredReferenceNode(null, table.ToString(), area, ColumnOrNull(firstColumn), ColumnOrNull(lastColumn));
            }

            public ValueNode ExternalStructureReference(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> table,
                StructuredReferenceArea area, ReadOnlySpan<char> firstColumn, ReadOnlySpan<char> lastColumn)
            {
                return new StructuredReferenceNode(new PrefixNode(new FileNode(workbookIndex), null, null, null), table.ToString(),
                    area, ColumnOrNull(firstColumn), ColumnOrNull(lastColumn));
            }

            public ValueNode Name(string context, SymbolRange range, ReadOnlySpan<char> name)
            {
                return new NameNode(null, name.ToString());
            }

            public ValueNode SheetName(string context, SymbolRange range, ReadOnlySpan<char> sheet, ReadOnlySpan<char> name)
            {
                var prefixNode = new PrefixNode(null, sheet.ToString(), null, null);
                return new NameNode(prefixNode, name.ToString());
            }

            public ValueNode BangName(string context, SymbolRange range, ReadOnlySpan<char> name)
            {
                return new NotSupportedNode("Bang name");
            }

            public ValueNode ExternalName(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> name)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), null, null, null);
                return new NameNode(prefixNode, name.ToString());
            }

            public ValueNode ExternalSheetName(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> sheet, ReadOnlySpan<char> name)
            {
                var prefixNode = new PrefixNode(new FileNode(workbookIndex), sheet.ToString(), null, null);
                return new NameNode(prefixNode, name.ToString());
            }

            public ValueNode BinaryNode(string context, SymbolRange range, BinaryOperation operation, ValueNode leftNode,
                ValueNode rightNode)
            {
                var op = operation switch
                {
                    BinaryOperation.Concat => BinaryOp.Concat,
                    BinaryOperation.GreaterOrEqualThan => BinaryOp.Gte,
                    BinaryOperation.LessOrEqualThan => BinaryOp.Lte,
                    BinaryOperation.LessThan => BinaryOp.Lt,
                    BinaryOperation.GreaterThan => BinaryOp.Gt,
                    BinaryOperation.NotEqual => BinaryOp.Neq,
                    BinaryOperation.Equal => BinaryOp.Eq,
                    BinaryOperation.Addition => BinaryOp.Add,
                    BinaryOperation.Subtraction => BinaryOp.Sub,
                    BinaryOperation.Multiplication => BinaryOp.Mult,
                    BinaryOperation.Division => BinaryOp.Div,
                    BinaryOperation.Power => BinaryOp.Exp,
                    BinaryOperation.Union => BinaryOp.Union,
                    BinaryOperation.Intersection => BinaryOp.Intersection,
                    BinaryOperation.Range => BinaryOp.Range,
                    _ => throw new NotSupportedException($"'{operation}' is not a binary operation.")
                };

                return new BinaryNode(op, leftNode, rightNode);
            }

            public ValueNode Unary(string context, SymbolRange range, UnaryOperation operation, ValueNode node)
            {
                var op = operation switch
                {
                    UnaryOperation.Plus => UnaryOp.Add,
                    UnaryOperation.Minus => UnaryOp.Subtract,
                    UnaryOperation.Percent => UnaryOp.Percentage,
                    UnaryOperation.ImplicitIntersection => UnaryOp.ImplicitIntersection,
                    UnaryOperation.SpillRange => UnaryOp.SpillRange,
                    _ => throw new NotSupportedException($"'{operation}' is not a unary operation.")
                };
                return new UnaryNode(op, node);
            }

            public ValueNode Nested(string context, SymbolRange range, ValueNode node)
            {
                return node;
            }

            private FunctionNode GetFunctionNode(PrefixNode? prefixNode, string functionName,
                IReadOnlyList<ValueNode> argumentNodes)
            {
                var foundFunction = _functionRegistry.TryGetFunc(functionName, out var minParams, out var maxParams);

                // If function is a future function, strip the prefix because all registration of functions
                // are without a prefix. That should change, but it's a reality for now.
                if (!foundFunction && functionName.StartsWith($"{DefaultFunctionNameSpace}."))
                {
                    functionName = functionName.Substring(DefaultFunctionNameSpace.Length + 1);
                    foundFunction = _functionRegistry.TryGetFunc(functionName, out minParams, out maxParams);
                }

                // Even if we haven't found anything, don't crash. Missing function will be evaluated to `#NAME?`
                if (!foundFunction)
                    return new FunctionNode(functionName, argumentNodes);

                var isWrongCount = (minParams != -1 && argumentNodes.Count < minParams) ||
                                   (maxParams != -1 && argumentNodes.Count > maxParams);
                if (isWrongCount && _argumentCount == ArgumentCount.Signal)
                    throw new WrongArgumentCountException();

                if (isWrongCount && _argumentCount == ArgumentCount.Ignore)
                    return new FunctionNode(prefixNode, functionName, argumentNodes);

                if (minParams != -1 && argumentNodes.Count < minParams)
                    throw new ExpressionParseException(
                        $"Too few parameters for function '{functionName}'. Expected a minimum of {minParams} and a maximum of {maxParams}.");

                if (maxParams != -1 && argumentNodes.Count > maxParams)
                    throw new ExpressionParseException(
                        $"Too many parameters for function '{functionName}'.Expected a minimum of {minParams} and a maximum of {maxParams}.");

                return new FunctionNode(prefixNode, functionName, argumentNodes);
            }

            /// <summary>
            /// The parser gives an empty column name for a reference to the whole table. The node uses <c>null</c> for it.
            /// </summary>
            private static string? ColumnOrNull(ReadOnlySpan<char> column)
            {
                return column.IsEmpty ? null : column.ToString();
            }

            private static XLError GetErrorValue(ReadOnlySpan<char> error)
            {
                if (!XLErrorParser.TryParseError(error.ToString(), out var errorEnum))
                    throw new InvalidOperationException($"'{error.ToString()}' is not error.");
                return errorEnum;
            }
        }
    }
}
