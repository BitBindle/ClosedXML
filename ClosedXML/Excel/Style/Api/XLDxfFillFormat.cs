using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// API object for dxf hidden behind IXLStyle.IXLFill interface.
/// </summary>
internal class XLDxfFillFormat : IXLFill
{
    private static readonly XLPatternFill _default = new()
    {
        PatternType = XLFillPatternValues.None,
        BackgroundColor = XLColor.Automatic,
        PatternColor = XLColor.Automatic
    };

    private readonly XLDxFormat _parent;

    internal XLDxfFillFormat(XLDxFormat parent)
    {
        _parent = parent;
    }

    XLColor IXLFill.BackgroundColor
    {
        get => Resolve(static fill => fill.Pattern?.BackgroundColor, _default.BackgroundColor);
        set => Modify(static (fill, bgColor) =>
        {
            var currentPattern = fill.Pattern ?? _default;
            var newPattern = currentPattern.WithModifiedBgColor(bgColor);
            return new XLDifferentialFillValue(newPattern);
        }, value);
    }

    XLColor IXLFill.PatternColor
    {
        get => Resolve(static fill => fill.Pattern?.PatternColor, _default.PatternColor);
        set => Modify(static (fill, patternColor) =>
        {
            var pattern = fill.Pattern ?? _default;
            return new XLDifferentialFillValue(pattern with { PatternColor = patternColor });
        }, value);
    }

    XLFillPatternValues IXLFill.PatternType
    {
        get => Resolve(static fill => fill.Pattern?.PatternType, _default.PatternType);
        set => Modify(static (fill, patternType) =>
        {
            var pattern = fill.Pattern ?? _default;
            var newPattern = pattern.WithModifiedPattern(patternType);
            return new XLDifferentialFillValue(newPattern);
        }, value);
    }

    IXLStyle IXLFill.SetBackgroundColor(XLColor value)
    {
        (this as IXLFill).BackgroundColor = value;
        return _parent;
    }

    IXLStyle IXLFill.SetPatternColor(XLColor value)
    {
        (this as IXLFill).PatternColor = value;
        return _parent;
    }

    IXLStyle IXLFill.SetPatternType(XLFillPatternValues value)
    {
        (this as IXLFill).PatternType = value;
        return _parent;
    }

    bool IEquatable<IXLFill>.Equals(IXLFill? other)
    {
        throw new NotSupportedException();
    }

    internal void SetValue(IXLFill value)
    {
        // The original should be valid and consistent.
        Modify(static (_, patternFill) => new XLDifferentialFillValue(new XLPatternFill
        {
            PatternType = patternFill.PatternType,
            PatternColor = patternFill.PatternColor,
            BackgroundColor = patternFill.BackgroundColor,
        }), value);
    }

    private T Resolve<T>(Func<XLDifferentialFillValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Fill, getProperty) ?? defaultValue;
    }

    private T Resolve<T>(Func<XLDifferentialFillValue, T?> getProperty, T defaultValue)
        where T : class
    {
        return _parent.Resolve(static format => format.Fill, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XLDifferentialFillValue, T, XLDifferentialFillValue> modify, T value)
    {
        _parent.ModifyFill(modify, value);
    }
}
