using System;
using ClosedXML.Excel.CalcEngine.Functions;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Logical
    {
        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("AND", 1, 255, And, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("FALSE", 0, 0, Adapt(False), FunctionFlags.Scalar);
            ce.RegisterFunction("IF", 2, 3, AdaptLastOptional(If, false), FunctionFlags.Range, AllowRange.Only, 1, 2);
            ce.RegisterFunction("IFERROR", 2, 2, Adapt(IfError), FunctionFlags.Scalar);
            ce.RegisterFunction("NOT", 1, 1, AdaptCoerced(Not), FunctionFlags.Scalar);
            ce.RegisterFunction("OR", 1, 255, Or, FunctionFlags.Range, AllowRange.All);
            ce.RegisterFunction("TRUE", 0, 0, Adapt(True), FunctionFlags.Scalar);
        }

        private static AnyValue And(CalcContext ctx, Span<AnyValue> args)
        {
            var aggResult = args.Aggregate(
                ctx,
                true,
                XLError.IncompatibleValue,
                static (acc, val) => acc && val,
                static (v, _) =>
                {
                    // Error or text other than "TRUE"/"FALSE" (e.g. "text") is returned as an error
                    if (!v.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var conversionError))
                        return conversionError;
                    return logical;
                },
                static v => v.IsLogical || v.IsNumber || v.IsError); // No text conversion for element of collection, blanks are ignored in references, errors are propagated

            if (!aggResult.TryPickT0(out var value, out var error))
                return error;

            return value;
        }

        private static ScalarValue False()
        {
            return false;
        }

        private static AnyValue If(ScalarValue condition, AnyValue valueIfTrue, AnyValue valueIfFalse)
        {
            if (!condition.TryCoerceLogicalOrBlankOrNumberOrText(out var value, out var error))
                return error;

            return value ? valueIfTrue : valueIfFalse;
        }

        private static AnyValue IfError(ScalarValue potentialError, ScalarValue alternative)
        {
            if (!potentialError.IsError)
                return potentialError.ToAnyValue();

            return alternative.ToAnyValue();
        }

        private static AnyValue Not(Boolean value)
        {
            return !value;
        }

        private static AnyValue Or(CalcContext ctx, Span<AnyValue> args)
        {
            var aggResult = args.Aggregate(
                ctx,
                false,
                XLError.IncompatibleValue,
                static (acc, val) => acc || val,
                static (v, _) =>
                {
                    // Error or text other than "TRUE"/"FALSE" (e.g. "text") is returned as an error
                    if (!v.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var conversionError))
                        return conversionError;
                    return logical;
                },
                static v => v.IsLogical || v.IsNumber || v.IsError); // No text conversion for element of collection, blanks are ignored in references, errors are propagated

            if (!aggResult.TryPickT0(out var value, out var error))
                return error;

            return value;
        }

        private static ScalarValue True()
        {
            return true;
        }
    }
}
