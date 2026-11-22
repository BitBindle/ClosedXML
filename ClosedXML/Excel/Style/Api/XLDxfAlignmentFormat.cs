using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

internal class XLDxfAlignmentFormat : IXLAlignment
{
    private readonly XLAlignmentFormatValue _default = XLAlignmentFormatValue.Default;
    private readonly XLDxFormat _parent;

    internal XLDxfAlignmentFormat(XLDxFormat parent)
    {
        _parent = parent;
    }

    XLAlignmentHorizontalValues IXLAlignment.Horizontal
    {
        get => Resolve(static alignment => alignment.Horizontal, _default.Horizontal);
        set => Modify(static (alignment, horizontal) => alignment with { Horizontal = horizontal }, value);
    }

    XLAlignmentVerticalValues IXLAlignment.Vertical
    {
        get => Resolve(static alignment => alignment.Vertical, _default.Vertical);
        set => Modify(static (alignment, vertical) => alignment with { Vertical = vertical }, value);
    }

    int IXLAlignment.Indent
    {
        get => Resolve(static alignment => alignment.Indent, _default.Indent);
        set => Modify(static (alignment, indent) => alignment with { Indent = indent }, value);
    }

    bool IXLAlignment.JustifyLastLine
    {
        get => Resolve(static alignment => alignment.JustifyLastLine, _default.JustifyLastLine);
        set => Modify(static (alignment, justifyLastLine) => alignment with { JustifyLastLine = justifyLastLine }, value);
    }

    XLAlignmentReadingOrderValues IXLAlignment.ReadingOrder
    {
        get => Resolve(static alignment => alignment.ReadingOrder, _default.ReadingOrder);
        set => Modify(static (alignment, readingOrder) => alignment with { ReadingOrder = readingOrder }, value);
    }

    int IXLAlignment.RelativeIndent
    {
        get => Resolve(static alignment => alignment.RelativeIndent, _default.RelativeIndent);
        set => Modify(static (alignment, relativeIndent) => alignment with { RelativeIndent = relativeIndent }, value);
    }

    bool IXLAlignment.ShrinkToFit
    {
        get => Resolve(static alignment => alignment.ShrinkToFit, _default.ShrinkToFit);
        set => Modify(static (alignment, shrinkToFit) => alignment with { ShrinkToFit = shrinkToFit }, value);
    }

    int IXLAlignment.TextRotation
    {
        get => Resolve(static alignment => alignment.TextRotation?.Value, _default.TextRotation.Value);
        set => Modify(static (alignment, textRotation) => alignment with { TextRotation = new TextRotation(textRotation) }, value);
    }

    bool IXLAlignment.WrapText
    {
        get => Resolve(static alignment => alignment.WrapText, _default.WrapText);
        set => Modify(static (alignment, wrapText) => alignment with { WrapText = wrapText }, value);
    }

    bool IXLAlignment.TopToBottom
    {
        get => Resolve(static alignment => alignment.TextRotation == TextRotation.VerticalText, false);
        set => Modify(static (alignment, topToBottom) => alignment with { TextRotation = topToBottom ? TextRotation.VerticalText : TextRotation.None }, value);
    }

    bool IEquatable<IXLAlignment>.Equals(IXLAlignment? other)
    {
        throw new NotSupportedException();
    }

    IXLStyle IXLAlignment.SetHorizontal(XLAlignmentHorizontalValues value)
    {
        (this as IXLAlignment).Horizontal = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetIndent(int value)
    {
        (this as IXLAlignment).Indent = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetJustifyLastLine()
    {
        return (this as IXLAlignment).SetJustifyLastLine(true);
    }

    IXLStyle IXLAlignment.SetJustifyLastLine(bool value)
    {
        (this as IXLAlignment).JustifyLastLine = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetReadingOrder(XLAlignmentReadingOrderValues value)
    {
        (this as IXLAlignment).ReadingOrder = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetRelativeIndent(int value)
    {
        (this as IXLAlignment).RelativeIndent = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetShrinkToFit()
    {
        return (this as IXLAlignment).SetShrinkToFit(true);
    }

    IXLStyle IXLAlignment.SetShrinkToFit(bool value)
    {
        (this as IXLAlignment).ShrinkToFit = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetTextRotation(int value)
    {
        (this as IXLAlignment).TextRotation = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetTopToBottom()
    {
        return (this as IXLAlignment).SetTopToBottom(true);
    }

    IXLStyle IXLAlignment.SetTopToBottom(bool value)
    {
        (this as IXLAlignment).TopToBottom = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetVertical(XLAlignmentVerticalValues value)
    {
        (this as IXLAlignment).Vertical = value;
        return _parent;
    }

    IXLStyle IXLAlignment.SetWrapText()
    {
        return (this as IXLAlignment).SetWrapText(true);
    }

    IXLStyle IXLAlignment.SetWrapText(bool value)
    {
        (this as IXLAlignment).WrapText = value;
        return _parent;
    }

    internal void SetValue(IXLAlignment value)
    {
        _parent.ModifyAlignment(static (alignment, value) => alignment with
        {
            Horizontal = value.Horizontal,
            Vertical = value.Vertical,
            TextRotation = new TextRotation(value.TextRotation),
            WrapText = value.WrapText,
            Indent = value.Indent,
            RelativeIndent = value.RelativeIndent,
            JustifyLastLine = value.JustifyLastLine,
            ShrinkToFit = value.ShrinkToFit,
            ReadingOrder = value.ReadingOrder,
        }, value);
    }

    private T Resolve<T>(Func<XLDifferentialAlignmentValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Alignment, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XLDifferentialAlignmentValue, T, XLDifferentialAlignmentValue> modify, T value)
    {
        _parent.ModifyAlignment(modify, value);
    }
}
