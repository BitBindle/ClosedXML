using System.Collections.Generic;
using System.Drawing;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel;

public class XLColorTests
{
    public static IEnumerable<object[]> VmlColors
    {
        get
        {
            // Hexadecimal color
            yield return new object[] { "#F0E0D0", Color.FromArgb(0xF0, 0xE0, 0xD0) };

            // Named color
            yield return new object[] { "red", Color.Red };

            // Palette color
            yield return new object[] { "Menu [30]", Color.FromArgb(0xF0, 0xF0, 0xF0) };
            yield return new object[] { "Menu", Color.FromArgb(0xF0, 0xF0, 0xF0) };

            // Unknown/malformed color
            yield return new object[] { "#NFOBACKGROUND", Color.FromName("#NFOBACKGROUND") };
        }
    }

    [TestCaseSource(nameof(VmlColors))]
    public void FromVmlColor_converts_hexadecimal_colors(string colorText, Color expectedColor)
    {
        var color = XLColor.FromVmlColor(colorText);

        Assert.That(color, Is.EqualTo(XLColor.FromColor(expectedColor)));
    }

    [Test]
    [Issue("60")]
    public void Equals_returns_false_for_null_and_other_types()
    {
        // IEquatable and object.Equals require false for null and for an object of a different type.
        var red = XLColor.Red;

        Assert.That(red.Equals((XLColor?)null), Is.False);
        Assert.That(red.Equals((object?)null), Is.False);
        Assert.That(red.Equals("Red"), Is.False);
        Assert.That(red.Equals((object)XLColor.FromColor(Color.Red)), Is.True);
    }
}
