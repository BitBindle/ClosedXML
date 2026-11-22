using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

// TODO Styles: Plan is to have a separate interface for dxf. Implement necessary features to pass tests, then replace.
/// <summary>
/// API object for dxf hidden behind IXLStyle.IXLBorder interface.
/// </summary>
internal class XLDxfBorderFormat : IXLBorder
{
    private static readonly XLBorderLine NoLine = XLBorderLine.None;
    private readonly XLDxFormat _parent;

    internal XLDxfBorderFormat(XLDxFormat parent)
    {
        _parent = parent;
    }

    XLBorderStyleValues IXLBorder.OutsideBorder
    {
        set => Modify(static (border, style) => border with
        {
            Outline = true,
            Left = (border.Left ?? NoLine) with { Style = style },
            Right = (border.Right ?? NoLine) with { Style = style },
            Top = (border.Top ?? NoLine) with { Style = style },
            Bottom = (border.Bottom ?? NoLine) with { Style = style },
        }, value);
    }

    XLColor IXLBorder.OutsideBorderColor
    {
        set => Modify(static (border, color) => border with
        {
            Outline = true,
            Left = (border.Left ?? NoLine) with { Color = color },
            Right = (border.Right ?? NoLine) with { Color = color },
            Top = (border.Top ?? NoLine) with { Color = color },
            Bottom = (border.Bottom ?? NoLine) with { Color = color },
        }, value);
    }

    XLBorderStyleValues IXLBorder.InsideBorder
    {
        set => Modify(static (border, style) => border with
        {
            Outline = false,
            Left = (border.Left ?? NoLine) with { Style = style },
            Right = (border.Right ?? NoLine) with { Style = style },
            Top = (border.Top ?? NoLine) with { Style = style },
            Bottom = (border.Bottom ?? NoLine) with { Style = style },
        }, value);
    }

    XLColor IXLBorder.InsideBorderColor
    {
        set => Modify(static (border, color) => border with
        {
            Outline = false,
            Left = (border.Left ?? NoLine) with { Color = color },
            Right = (border.Right ?? NoLine) with { Color = color },
            Top = (border.Top ?? NoLine) with { Color = color },
            Bottom = (border.Bottom ?? NoLine) with { Color = color },
        }, value);
    }

    XLBorderStyleValues IXLBorder.LeftBorder
    {
        get => Resolve(static x => x.Left?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Left = (border.Left ?? NoLine) with { Style = style } }, value);
    }

    XLColor IXLBorder.LeftBorderColor
    {
        get => Resolve(static border => border.Left?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Left = (border.Left ?? NoLine) with { Color = color } }, value);
    }

    XLBorderStyleValues IXLBorder.RightBorder
    {
        get => Resolve(static x => x.Right?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Right = (border.Right ?? NoLine) with { Style = style } }, value);
    }

    XLColor IXLBorder.RightBorderColor
    {
        get => Resolve(static border => border.Right?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Right = (border.Right ?? NoLine) with { Color = color } }, value);
    }

    XLBorderStyleValues IXLBorder.TopBorder
    {
        get => Resolve(static x => x.Top?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Top = (border.Top ?? NoLine) with { Style = style } }, value);
    }

    XLColor IXLBorder.TopBorderColor
    {
        get => Resolve(static border => border.Top?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Top = (border.Top ?? NoLine) with { Color = color } }, value);
    }

    XLBorderStyleValues IXLBorder.BottomBorder
    {
        get => Resolve(static x => x.Bottom?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Bottom = (border.Bottom ?? NoLine) with { Style = style } }, value);
    }

    XLColor IXLBorder.BottomBorderColor
    {
        get => Resolve(static border => border.Bottom?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Bottom = (border.Bottom ?? NoLine) with { Color = color } }, value);
    }

    bool IXLBorder.DiagonalUp
    {
        get => Resolve(static border => (bool?)border.DiagonalUp, false);
        set => Modify(static (border, diagonalUp) => border with { DiagonalUp = diagonalUp }, value);
    }

    bool IXLBorder.DiagonalDown
    {
        get => Resolve(static border => (bool?)border.DiagonalDown, false);
        set => Modify(static (border, diagonalDown) => border with { DiagonalDown = diagonalDown }, value);
    }

    XLBorderStyleValues IXLBorder.DiagonalBorder
    {
        get => Resolve(static x => x.Diagonal?.Style, NoLine.Style);
        set => Modify(static (border, style) => border with { Diagonal = (border.Diagonal ?? NoLine) with { Style = style } }, value);
    }

    XLColor IXLBorder.DiagonalBorderColor
    {
        get => Resolve(static border => border.Diagonal?.Color, NoLine.Color);
        set => Modify(static (border, color) => border with { Diagonal = (border.Diagonal ?? NoLine) with { Color = color } }, value);
    }

    internal void SetValue(IXLBorder value)
    {
        Modify((border, other) => border with
        {
            Left = new XLBorderLine(other.LeftBorderColor, other.LeftBorder),
            Right = new XLBorderLine(other.RightBorderColor, other.RightBorder),
            Top = new XLBorderLine(other.TopBorderColor, other.TopBorder),
            Bottom = new XLBorderLine(other.BottomBorderColor, other.BottomBorder),
            DiagonalUp = other.DiagonalUp,
            DiagonalDown = other.DiagonalDown,
            Diagonal = new XLBorderLine(other.DiagonalBorderColor, other.DiagonalBorder),
            Horizontal = null,
            Vertical = null,
            Outline = true,
        }, value);
    }

    private T Resolve<T>(Func<XLDifferentialBorderValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Border, getProperty) ?? defaultValue;
    }

    private T Resolve<T>(Func<XLDifferentialBorderValue, T?> getProperty, T defaultValue)
        where T : class
    {
        return _parent.Resolve(static format => format.Border, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XLDifferentialBorderValue, T, XLDifferentialBorderValue> modify, T value)
    {
        _parent.ModifyBorder(modify, value);
    }

    IXLStyle IXLBorder.SetOutsideBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).OutsideBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetOutsideBorderColor(XLColor value)
    {
        (this as IXLBorder).OutsideBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetInsideBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).InsideBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetInsideBorderColor(XLColor value)
    {
        (this as IXLBorder).InsideBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetLeftBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).LeftBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetLeftBorderColor(XLColor value)
    {
        (this as IXLBorder).LeftBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetRightBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).RightBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetRightBorderColor(XLColor value)
    {
        (this as IXLBorder).RightBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetTopBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).TopBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetTopBorderColor(XLColor value)
    {
        (this as IXLBorder).TopBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetBottomBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).BottomBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetBottomBorderColor(XLColor value)
    {
        (this as IXLBorder).BottomBorderColor = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalUp()
    {
        return (this as IXLBorder).SetDiagonalUp(true);
    }

    IXLStyle IXLBorder.SetDiagonalUp(bool value)
    {
        (this as IXLBorder).DiagonalUp = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalDown()
    {
        return (this as IXLBorder).SetDiagonalDown(true);
    }

    IXLStyle IXLBorder.SetDiagonalDown(bool value)
    {
        (this as IXLBorder).DiagonalDown = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).DiagonalBorder = value;
        return _parent;
    }

    IXLStyle IXLBorder.SetDiagonalBorderColor(XLColor value)
    {
        (this as IXLBorder).DiagonalBorderColor = value;
        return _parent;
    }

    bool IEquatable<IXLBorder>.Equals(IXLBorder? other)
    {
        throw new NotImplementedException();
    }
}
